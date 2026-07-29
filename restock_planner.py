"""
Next-Month Restock & Open-To-Buy (OTB) Planning System
=========================================================
Reuses finance_analyzer.py's loaders (SKU cost master, bank statement, D2C/
dropship revenue) so cost and cash figures stay identical to the dashboard --
never a second, drifting copy of the same numbers.

Inputs (all real, none fabricated -- see each loader's docstring for exactly
where a number comes from):
  active-isku-master.xlsx            -- Bun's manually-curated list of which
                                         SKUs are currently orderable, grouped
                                         into 14 supplier/brand "Order Sheet"
                                         sections. Only SKUs on this list are
                                         ever recommended -- anything not on
                                         it is ignored, per Bun's own rule.
  inventory-forecast-SiteGiant/
    last-30days-*.xlsx                -- SiteGiant's own Inventory Forecasting
                                          export, rolling 30-day window as of
                                          export date. Primary source for
                                          velocity (sale_per_day), current
                                          stock, and days-of-stock-left.
    2026-0[3-6]-*.xlsx                -- four calendar-month exports (trend
                                          context only -- stock_on_hand in
                                          these is TODAY's stock, not a true
                                          historical snapshot, so they're never
                                          used for the restock math itself).
    all-time-*.xlsx                   -- lifetime sales per SKU (confirmed by
                                          Bun 2026-07-21), used only to flag
                                          "never sold" vs. "used to sell, now
                                          stalled" -- not a budget input.
  isku_database.json -- via finance_analyzer's own
                                          load_sku_master()/apply_cost_overrides(),
                                          so restock cost-per-unit is exactly
                                          the same number the dashboard uses.
  e-Statement CSV (OCBC)              -- via finance_analyzer's own
                                          load_bank_statement(), for the real
                                          current cash position.

Output: Next_Month_Restock_Proposal.md
"""
import json
import math
from pathlib import Path

import pandas as pd

import finance_analyzer as fa

BASE_DIR = fa.BASE_DIR
FORECAST_DIR = BASE_DIR / "inventory-forecast-SiteGiant"
ACTIVE_ISKU_PATH = BASE_DIR / "active-isku-master.xlsx"

# Per-class safety-stock target (days of forward cover to restock up to).
# A-class (top ~80% of revenue): never stock out -- 30 days.
# B-class (next ~15%): 15 days, per Bun's Gemini brief.
# C-class (bottom ~5%, slow movers): 7 days only -- standard retail practice
# is NOT to proactively replenish slow movers to a full safety-stock level;
# only top them up enough to avoid a hard stockout on an active listing.
CLASS_TARGET_DAYS = {"A": 30, "B": 15, "C": 7}

# Below this many days of stock left, a SKU is flagged urgent regardless of class.
URGENT_DIL_THRESHOLD = 20

# Hard ceiling on total restock spend as a share of last month's actual
# revenue -- protects cash flow from an overly aggressive recommendation.
# This is Bun's own ceiling (from the Gemini brief); NOT a forecast of what
# revenue will be next month, just a cap on THIS month's known number.
REVENUE_CAP_PCT = 0.40


# ---------------------------------------------------------------------------
# 1. Active ISKU catalog -- dynamically parsed so it stays correct as Bun
# edits the sheet (adds/removes SKUs, per his own note in the file: "when
# some isku need to remove or some isku need to add in new, i will inform
# you after and please help to update"). Never hardcode row numbers.
# ---------------------------------------------------------------------------
def load_active_isku_catalog():
    """The file also contains a short-form table of contents (just brand
    names like "Order Sheet-SL" with no data under them) in two places --
    once near the top, once duplicated near the bottom, both of which also
    start with "Order Sheet" and would otherwise be mistaken for real
    section headers. Verified 2026-07-21: a real section header is ALWAYS
    followed immediately by the column-title row ("SKU","EAN Code",...) --
    the TOC entries aren't. That's the discriminator used below, not row
    position (row numbers shift as Bun edits this file over time).
    """
    if not ACTIVE_ISKU_PATH.exists():
        raise FileNotFoundError(f"Could not find {ACTIVE_ISKU_PATH.name} in the project root.")
    wb = fa.load_workbook(ACTIVE_ISKU_PATH, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))

    def is_real_section_header(i):
        r = rows[i]
        if not (r and r[0] and isinstance(r[0], str) and r[0].strip().startswith("Order Sheet")):
            return False
        nxt = rows[i + 1] if i + 1 < len(rows) else None
        return bool(nxt and nxt[0] and str(nxt[0]).strip() == "SKU")

    section_idxs = [i for i in range(len(rows)) if is_real_section_header(i)]
    if not section_idxs:
        raise ValueError(f"{ACTIVE_ISKU_PATH.name}: no 'Order Sheet-...' section markers found -- "
                          f"format may have changed, refusing to guess.")
    section_idxs.append(len(rows))  # end boundary for the last section

    records = []
    for n in range(len(section_idxs) - 1):
        start, end = section_idxs[n], section_idxs[n + 1]
        order_sheet = rows[start][0].strip()
        # start = section title row, start+1 = column-header row (SKU/EAN
        # Code/Series/Brand/Model/Color/Quantity/Price), data begins at
        # start+2. A real product row always carries a Price (col index 7);
        # free-text notes that sometimes appear inside/after a section (e.g.
        # "above are all the active listing...") only ever populate column A,
        # so requiring a Price filters them out without needing to locate an
        # exact blank-row boundary.
        for r in rows[start + 2:end]:
            if not r or r[0] in (None, ""):
                continue
            if len(r) <= 7 or r[7] is None:
                continue
            records.append({
                "order_sheet": order_sheet, "sku": str(r[0]).strip(),
                "ean_code": r[1] if len(r) > 1 else None,
                "series": r[2] if len(r) > 2 else None,
                "brand": r[3] if len(r) > 3 else None,
                "model": r[4] if len(r) > 4 else None,
                "color": r[5] if len(r) > 5 else None,
            })
    df = pd.DataFrame(records).drop_duplicates(subset="sku", keep="first")
    fa.warn(f"Active ISKU catalog: {len(df)} active SKUs loaded across "
             f"{df['order_sheet'].nunique()} order sheets from {ACTIVE_ISKU_PATH.name}.")
    return df


# ---------------------------------------------------------------------------
# 2. SiteGiant Inventory Forecasting exports
# ---------------------------------------------------------------------------
def load_forecast_file(path):
    """One SiteGiant 'Inventory Forecasting' export. '---' means SiteGiant
    itself couldn't compute a days-of-stock figure (e.g. zero velocity) --
    parsed as NaN, not 0, so it's never mistaken for "0 days left / urgent"."""
    df = pd.read_excel(path, sheet_name=0)
    df = df.rename(columns={
        "estimated_stock_without_purchase_order (day)": "days_left_no_po",
        "estimated_stock_with_purchase_order (day)": "days_left_with_po",
    })
    numeric_cols = ["stock_on_hand", "stock_on_purchase_order", "safety_stock", "total_sales",
                     "sale_per_day", "lead_time", "recommended_quantity", "days_left_no_po", "days_left_with_po"]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].replace("---", pd.NA), errors="coerce")
    df["isku"] = df["isku"].astype(str).str.strip()
    return df.drop_duplicates(subset="isku", keep="first")


def load_all_forecast_files():
    if not FORECAST_DIR.exists():
        raise FileNotFoundError(f"Could not find {FORECAST_DIR.name}/ in the project root.")
    files = {p.name: p for p in FORECAST_DIR.glob("*.xlsx") if not p.name.startswith("~$")}

    last30_matches = [p for n, p in files.items() if n.startswith("last-30days")]
    alltime_matches = [p for n, p in files.items() if n.startswith("all-time")]
    monthly_matches = sorted(p for n, p in files.items() if n[:7].replace("-", "").isdigit() and n.startswith("2026-"))

    if not last30_matches:
        raise FileNotFoundError(f"No 'last-30days-*.xlsx' file in {FORECAST_DIR.name}/ -- "
                                 f"this is the primary velocity source, required.")
    last30 = load_forecast_file(last30_matches[0])
    alltime = load_forecast_file(alltime_matches[0]) if alltime_matches else None
    monthly = {p.stem[:7]: load_forecast_file(p) for p in monthly_matches}
    return last30, alltime, monthly


# ---------------------------------------------------------------------------
# 3. ABC classification -- standard Pareto cut on 30-day revenue, computed
# only across active SKUs with a known sell price. A SKU with real velocity
# but no RSP/DSP price on file gets flagged "class n/a" rather than silently
# defaulted into a bucket using a fabricated price.
# ---------------------------------------------------------------------------
def classify_abc(df):
    df = df.copy()
    priced = df[df["revenue_30d"].notna()].sort_values("revenue_30d", ascending=False)
    total_rev = priced["revenue_30d"].sum()
    df["abc_class"] = None
    if total_rev > 0:
        priced["cum_pct"] = priced["revenue_30d"].cumsum() / total_rev * 100
        priced["abc_class"] = priced["cum_pct"].apply(lambda p: "A" if p <= 80 else ("B" if p <= 95 else "C"))
        df.loc[priced.index, "abc_class"] = priced["abc_class"]
    return df


# ---------------------------------------------------------------------------
# 4. Real cash position + last month's actual Revenue/Payroll/Bank OpEx --
# reuse finance_analyzer's own pipeline so these never drift from the
# dashboard's numbers. Only the pieces needed here are computed.
# ---------------------------------------------------------------------------
def load_finance_context():
    sku_master = fa.apply_cost_overrides(fa.load_sku_master())
    orders = fa.load_sitegiant_orders()
    sitegiant = fa.sitegiant_channel_metrics(orders, sku_master)
    dropship = fa.dropship_channel_metrics(sku_master)
    bank_df = fa.load_bank_statement()
    bank_opex_result = fa.bank_opex(bank_df)
    payroll_df = fa.load_payroll()
    payroll_monthly_df = fa.payroll_monthly(payroll_df)
    target_month = (sitegiant["monthly"]["month"].iloc[0]
                     if not sitegiant["monthly"].empty else pd.Period(fa.datetime.now(), freq="M"))
    platform_fees = fa.load_platform_fees(orders, target_month)
    net_profit = fa.compute_net_profit(sitegiant, dropship, bank_opex_result, payroll_monthly_df,
                                        target_month, platform_fees)

    bank_df_sorted = bank_df.sort_values("Statement Date")
    cash_position = float(bank_df_sorted["Closing Book Balance"].iloc[-1]) if "Closing Book Balance" in bank_df_sorted.columns else None
    as_of = bank_df_sorted["Statement Date"].iloc[-1]

    return {
        "sku_master": sku_master, "target_month": str(target_month),
        "revenue": net_profit["revenue"], "payroll_cost": net_profit["payroll_cost"],
        "bank_opex": net_profit["bank_opex"], "cash_position": cash_position,
        "cash_position_as_of": as_of.strftime("%Y-%m-%d") if pd.notna(as_of) else None,
    }


# ---------------------------------------------------------------------------
# 5. Core recommendation build
# ---------------------------------------------------------------------------
def build_recommendations(active_df, last30, alltime, sku_master):
    df = active_df.merge(last30, left_on="sku", right_on="isku", how="left", indicator=True)

    not_in_forecast = df[df["_merge"] == "left_only"][["sku", "order_sheet", "brand"]]
    df = df[df["_merge"] == "both"].drop(columns=["_merge", "isku"])

    if alltime is not None:
        df = df.merge(alltime[["isku", "total_sales"]].rename(columns={"isku": "sku", "total_sales": "lifetime_sales"}),
                       on="sku", how="left")
    else:
        df["lifetime_sales"] = pd.NA

    df = df.merge(sku_master[["cost", "rsp", "dsp"]], left_on="sku", right_index=True, how="left")
    no_cost = df[df["cost"].isna()][["sku", "order_sheet", "brand", "stock_on_hand", "sale_per_day"]]
    df = df[df["cost"].notna()].copy()

    # Revenue proxy for ABC ranking: 30-day sold qty x sell price (RSP,
    # falling back to DSP if RSP isn't set -- both are real listed prices
    # from the cost master, never a guessed markup).
    df["sell_price"] = df["rsp"].where(df["rsp"].notna() & (df["rsp"] > 0), df["dsp"])
    df["revenue_30d"] = df["total_sales"] * df["sell_price"]
    df = classify_abc(df)

    df["days_left"] = df["days_left_no_po"]  # SiteGiant's own calc, NaN = undefined (no recent sales)
    df["is_oversold"] = df["stock_on_hand"] < 0

    def target_days(row):
        if row["is_oversold"]:
            return CLASS_TARGET_DAYS.get(row["abc_class"], CLASS_TARGET_DAYS["C"])
        return CLASS_TARGET_DAYS.get(row["abc_class"]) if pd.notna(row["abc_class"]) else None

    df["target_days"] = df.apply(target_days, axis=1)

    def recommended_qty(row):
        # No class (no price on file) and no oversold correction needed --
        # can still flag urgency from days_left, but there's no safety-stock
        # target to size a quantity against, so recommend 0 rather than guess.
        if row["target_days"] is None:
            return 0.0 if not row["is_oversold"] else float(-row["stock_on_hand"])
        need = row["target_days"] * (row["sale_per_day"] or 0) - row["stock_on_hand"] - row["stock_on_purchase_order"]
        return max(0.0, math.ceil(need)) if pd.notna(need) else 0.0

    df["recommended_qty"] = df.apply(recommended_qty, axis=1)
    df["est_cost"] = df["recommended_qty"] * df["cost"]
    df["urgent"] = df["is_oversold"] | (df["days_left"].notna() & (df["days_left"] < URGENT_DIL_THRESHOLD))

    return df, not_in_forecast, no_cost


# ---------------------------------------------------------------------------
# 6. Open-To-Buy budget cap -- the smallest of: what's actually needed, 40%
# of last month's real revenue, and (real cash on hand minus one month of
# real Payroll+Bank OpEx as an untouchable float). The cash constraint is
# grounded in actual figures already in the system, not a guessed percentage.
# ---------------------------------------------------------------------------
def apply_otb_cap(reco_df, ctx):
    raw_need = float(reco_df["est_cost"].sum())
    revenue_cap = ctx["revenue"] * REVENUE_CAP_PCT
    cash_reserve = ctx["payroll_cost"] + ctx["bank_opex"]
    cash_cap = max(0.0, (ctx["cash_position"] or 0) - cash_reserve) if ctx["cash_position"] is not None else float("inf")

    final_budget = min(raw_need, revenue_cap, cash_cap)
    binding = "none (full recommended amount fits)"
    if final_budget == cash_cap < min(raw_need, revenue_cap):
        binding = "cash position (after reserving last month's Payroll + Bank OpEx)"
    elif final_budget == revenue_cap < raw_need:
        binding = f"{REVENUE_CAP_PCT*100:.0f}% of last month's revenue ceiling"

    # Priority order when trimming to fit: oversold corrections first, then
    # urgency (fewest days left first) within each class, C last.
    class_rank = {"A": 0, "B": 1, "C": 2}
    reco_df = reco_df.copy()
    reco_df["_priority"] = list(zip(
        (~reco_df["is_oversold"]),  # oversold (False) sorts before not-oversold (True)
        reco_df["abc_class"].map(class_rank).fillna(3),
        reco_df["days_left"].fillna(9999),
    ))
    reco_df = reco_df.sort_values("_priority")
    reco_df["cum_cost"] = reco_df["est_cost"].cumsum()
    reco_df["funded"] = reco_df["cum_cost"] <= final_budget + 1e-9
    reco_df = reco_df.drop(columns=["_priority"])

    return reco_df, {
        "raw_need": raw_need, "revenue_cap": revenue_cap, "cash_cap": cash_cap,
        "final_budget": final_budget, "binding_constraint": binding,
        "cash_reserve": cash_reserve,
    }


# ---------------------------------------------------------------------------
# 7. Report
# ---------------------------------------------------------------------------
def write_report(reco_df, budget, not_in_forecast, no_cost, ctx, out_path):
    L = []
    L.append("# Next-Month Restock Proposal\n")
    L.append(f"*Generated from {ctx['target_month']} actuals + SiteGiant Inventory Forecasting "
             f"(last 30 days as of the export date) + cash position as of {ctx['cash_position_as_of']}.*\n")

    funded = reco_df[reco_df["funded"] & (reco_df["recommended_qty"] > 0)]
    cut = reco_df[~reco_df["funded"] & (reco_df["recommended_qty"] > 0)]
    urgent_unfunded = cut[cut["urgent"]]

    L.append("## 1. Budget Summary\n")
    L.append("| | Amount |")
    L.append("|---|---:|")
    L.append(f"| Raw restock need (all flagged active SKUs) | {fa.fmt_money(budget['raw_need'])} |")
    L.append(f"| Ceiling: {REVENUE_CAP_PCT*100:.0f}% of last month's revenue ({fa.fmt_money(ctx['revenue'])}) | {fa.fmt_money(budget['revenue_cap'])} |")
    L.append(f"| Ceiling: cash on hand ({fa.fmt_money(ctx['cash_position'])}) minus last month's Payroll+Bank OpEx reserve ({fa.fmt_money(budget['cash_reserve'])}) | {fa.fmt_money(budget['cash_cap'])} |")
    L.append(f"| **Recommended budget** | **{fa.fmt_money(budget['final_budget'])}** |")
    L.append(f"| Binding constraint | {budget['binding_constraint']} |\n")

    if len(cut):
        L.append(f"**{len(cut)} SKU(s) totalling {fa.fmt_money(cut['est_cost'].sum())} were cut to fit the budget** "
                  f"(lowest priority: C-class and/or most days-of-stock-left trimmed first). "
                  f"{len(urgent_unfunded)} of those are still flagged **urgent** (< {URGENT_DIL_THRESHOLD} days "
                  f"stock or already oversold) -- reviewed at the bottom of this report.\n")

    L.append("## 2. Recommended Restock Orders (within budget)\n")
    L.append(f"{len(funded)} SKUs, {fa.fmt_money(funded['est_cost'].sum())} total.\n")
    L.append("| SKU | Order Sheet | Class | Stock | On Order | Sale/Day | Days Left | Recommend Qty | Unit Cost | Est. Cost |")
    L.append("|---|---|---|---:|---:|---:|---:|---:|---:|---:|")
    for _, r in funded.sort_values("est_cost", ascending=False).iterrows():
        days_left = f"{r['days_left']:.0f}" if pd.notna(r["days_left"]) else "n/a"
        L.append(f"| {r['sku']} | {r['order_sheet']} | {r['abc_class'] or 'n/a'} | {r['stock_on_hand']:.0f} | "
                  f"{r['stock_on_purchase_order']:.0f} | {r['sale_per_day']:.2f} | {days_left} | "
                  f"{r['recommended_qty']:.0f} | {fa.fmt_money(r['cost'])} | {fa.fmt_money(r['est_cost'])} |")
    L.append("")

    if len(cut):
        L.append("## 3. Cut for Budget (not funded this round)\n")
        L.append("| SKU | Order Sheet | Class | Days Left | Recommend Qty | Est. Cost | Urgent? |")
        L.append("|---|---|---|---:|---:|---:|---|")
        for _, r in cut.sort_values(["urgent", "est_cost"], ascending=[False, False]).iterrows():
            days_left = f"{r['days_left']:.0f}" if pd.notna(r["days_left"]) else "n/a"
            L.append(f"| {r['sku']} | {r['order_sheet']} | {r['abc_class'] or 'n/a'} | {days_left} | "
                      f"{r['recommended_qty']:.0f} | {fa.fmt_money(r['est_cost'])} | {'YES' if r['urgent'] else ''} |")
        L.append("")

    L.append("## 4. Data Gaps -- Not Included Above (need manual attention)\n")
    if len(no_cost):
        L.append(f"**{len(no_cost)} active SKU(s) have no cost in the master table** -- can't size a budget "
                  f"without a real cost, so excluded rather than guessed. Same SKUs will show up in the "
                  f"dashboard's Pending Action table once you check the finance system.\n")
        L.append("| SKU | Order Sheet | Stock | Sale/Day |")
        L.append("|---|---|---:|---:|")
        for _, r in no_cost.iterrows():
            L.append(f"| {r['sku']} | {r['order_sheet']} | {r['stock_on_hand']:.0f} | {r['sale_per_day']:.2f} |")
        L.append("")
    if len(not_in_forecast):
        L.append(f"**{len(not_in_forecast)} active SKU(s) are on your active-isku-master list but weren't found "
                  f"in the SiteGiant forecast export** -- either not set up in SiteGiant yet (e.g. JJT SKUs, "
                  f"per your own note in the file) or a SKU-code mismatch. Review manually.\n")
        L.append("| SKU | Order Sheet | Brand |")
        L.append("|---|---|---|")
        for _, r in not_in_forecast.iterrows():
            L.append(f"| {r['sku']} | {r['order_sheet']} | {r['brand']} |")
        L.append("")

    L.append("## Methodology\n")
    L.append(f"- **Velocity/stock**: SiteGiant's own last-30-day Inventory Forecasting export -- `sale_per_day`, "
              f"`stock_on_hand`, `stock_on_purchase_order`, and days-of-stock-left are all SiteGiant's own "
              f"numbers, not recomputed.\n"
              f"- **ABC class**: standard Pareto cut on 30-day revenue (qty sold x RSP/DSP price) -- A = top 80% "
              f"cumulative revenue, B = next 15%, C = bottom 5%. SKUs with real sales but no listed price are "
              f"marked class n/a rather than guessed into a bucket.\n"
              f"- **Target restock level**: A=30 days cover, B=15 days, C=7 days (slow movers are topped up to "
              f"avoid a hard stockout, not fully replenished -- standard retail practice). Oversold SKUs "
              f"(negative stock) are corrected regardless of class.\n"
              f"- **Budget cap**: the smallest of (a) what's actually needed, (b) {REVENUE_CAP_PCT*100:.0f}% of "
              f"last month's real revenue, (c) real cash on hand minus a reserve equal to last month's real "
              f"Payroll + Bank OpEx. Nothing here is a forecast of next month -- all three inputs are actuals.\n"
              f"- SiteGiant's own `recommended_quantity`, `safety_stock`, and `lead_time` fields are populated "
              f"for only a handful of SKUs in your account (not configured), so this report computes its own "
              f"figures from `sale_per_day` and stock levels instead of relying on those columns.\n")

    out_path.write_text("\n".join(L), encoding="utf-8")
    return out_path


# ---------------------------------------------------------------------------
# 8. JSON handoff -- consumed by finance_analyzer.py's dashboard as an
# optional extra tab (additive: dashboard.html builds fine without this file,
# same pattern as the platform-fee/income-report inputs there). Kept as a
# separate file/pipeline rather than merged into finance_analyzer.main()
# because the two run on different cadences -- financial numbers get
# refreshed far more often than a fresh SiteGiant forecast export.
# ---------------------------------------------------------------------------
def _clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    return str(v)


def _row_records(df, cols):
    return [{c: _clean(r[c]) for c in cols} for _, r in df.iterrows()]


def write_restock_json(reco_df, budget, not_in_forecast, no_cost, ctx, out_path):
    funded = reco_df[reco_df["funded"] & (reco_df["recommended_qty"] > 0)].sort_values("est_cost", ascending=False)
    cut = reco_df[~reco_df["funded"] & (reco_df["recommended_qty"] > 0)].sort_values("est_cost", ascending=False)

    order_cols = ["sku", "order_sheet", "abc_class", "stock_on_hand", "stock_on_purchase_order",
                  "sale_per_day", "days_left", "recommended_qty", "cost", "est_cost", "urgent"]

    data = {
        "generated_at": fa.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "target_month": ctx["target_month"],
        "cash_position": _clean(ctx["cash_position"]),
        "cash_position_as_of": ctx["cash_position_as_of"],
        "budget": {
            "raw_need": _clean(budget["raw_need"]), "revenue_cap": _clean(budget["revenue_cap"]),
            "cash_cap": _clean(budget["cash_cap"]), "final_budget": _clean(budget["final_budget"]),
            "binding_constraint": budget["binding_constraint"], "cash_reserve": _clean(budget["cash_reserve"]),
            "revenue_cap_pct": REVENUE_CAP_PCT * 100, "last_month_revenue": _clean(ctx["revenue"]),
        },
        "funded": _row_records(funded, order_cols),
        "cut": _row_records(cut, order_cols),
        "no_cost": _row_records(no_cost, ["sku", "order_sheet", "stock_on_hand", "sale_per_day"]),
        "not_in_forecast": _row_records(not_in_forecast, ["sku", "order_sheet", "brand"]),
        "summary": {
            "active_skus_matched": int(len(reco_df)),
            "funded_count": int(len(funded)), "funded_spend": _clean(funded["est_cost"].sum()),
            "cut_count": int(len(cut)), "cut_spend": _clean(cut["est_cost"].sum()),
            "urgent_unfunded_count": int(cut["urgent"].sum()) if len(cut) else 0,
        },
    }
    out_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return out_path


def main():
    print("Loading active ISKU catalog...")
    active_df = load_active_isku_catalog()

    print("Loading SiteGiant Inventory Forecasting exports...")
    last30, alltime, monthly = load_all_forecast_files()

    print("Loading finance context (cost master, revenue, payroll, bank cash position)...")
    ctx = load_finance_context()

    print("Building restock recommendations...")
    reco_df, not_in_forecast, no_cost = build_recommendations(active_df, last30, alltime, ctx["sku_master"])

    print("Applying Open-To-Buy budget cap...")
    reco_df, budget = apply_otb_cap(reco_df, ctx)

    out_path = BASE_DIR / "Next_Month_Restock_Proposal.md"
    write_report(reco_df, budget, not_in_forecast, no_cost, ctx, out_path)

    json_path = BASE_DIR / "restock_data.json"
    write_restock_json(reco_df, budget, not_in_forecast, no_cost, ctx, json_path)

    funded = reco_df[reco_df["funded"] & (reco_df["recommended_qty"] > 0)]
    print(f"\n=== SUMMARY ===")
    print(f"Active SKUs matched to forecast data: {len(reco_df)}")
    print(f"Recommended budget: {fa.fmt_money(budget['final_budget'])} ({budget['binding_constraint']})")
    print(f"SKUs funded: {len(funded)} / spend {fa.fmt_money(funded['est_cost'].sum())}")
    print(f"\nReport written to: {out_path}")
    print(f"Dashboard data written to: {json_path} -- re-run finance_analyzer.py to fold it into dashboard.html")
    if fa.WARNINGS:
        print(f"{len(fa.WARNINGS)} data-quality warnings logged during load.")


if __name__ == "__main__":
    main()
