"""
One-time (but re-runnable) migration: build isku_database.json (the local
source of truth for ISKU pricing/costing) from the Google Sheet CSV exports +
the existing cost_overrides.json patch file. Writes schema_version 2 directly
(currency decoupled from ITEC_CODE, ITEC_CODE trimmed to Bun's 3 active China
presets with an itemized invisible-costing fees list, per his 2026-07-23
direction). Re-run this for a full from-scratch rebuild off the sheet CSVs,
e.g. after a bulk re-export -- it overwrites isku_database.json entirely, so
don't run it if you've since made manual edits through the ISKU Manager tab
that aren't reflected in a fresh sheet export.

Run once from this folder:
    python migrate_isku_database.py

Reads:
    cogs-reference-google-sheet-database_<latest date>.csv   -- main Database tab export
    database-dsp-pp_<latest date>.csv                        -- offline/physical-partner pricing tab export
    cost_overrides.json                                      -- existing manual cost patches (folded in, then can be retired)

Writes:
    isku_database.json
"""
import glob
import json
import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

# The OLD ITEC_CODE set as it exists in the raw sheet CSV export -- used only
# to figure out what currency each historical row was actually priced in
# (schema v1 bundled currency into ITEC_CODE; v2 splits them apart). Not
# written to isku_database.json.
OLD_CODE_CURRENCY = {
    "IT-CN-S": "CNY", "IT-CN-B": "CNY", "IT-CN-SS": "CNY",
    "IT-CN-S FRESCO": "CNY", "IT-CN-S FRESCO CARD": "CNY",
    "IT-MLY": "MYR",
    "IT-TW": "USD", "IT-CN(US)": "USD", "IT-KOR": "USD",
}

# RM per 1 unit of foreign currency. CNY/USD are the same real rates already
# in the sheet (old divisor form: CNY 1.60, USD 0.21 -- "X foreign per 1 RM"),
# just algebraically flipped to a multiplier ("Y RM per 1 foreign"). SGD is
# Bun's own already-confirmed rate from elsewhere in this project (2026-07-21):
# 1 SGD = 3.0 MYR.
NEW_CURRENCIES = {
    "MYR": {"rate_to_rm": 1.0},
    "CNY": {"rate_to_rm": round(1 / 1.60, 4)},
    "USD": {"rate_to_rm": round(1 / 0.21, 4)},
    "SGD": {"rate_to_rm": 3.0},
}

# Trimmed to the 3 China presets Bun actually uses (2026-07-23 direction) --
# no longer currency/tax lookups, just his itemized "invisible costing" per
# item (packaging/shipping, engraving, custom packaging, etc.).
NEW_ITEC_CODES = {
    "IT-CN-B": {
        "description": "China -- biggest size item",
        "tax_rate": 0.0,
        "fees": [{"label": "Packaging/Shipping", "amount": 2.0}],
    },
    "IT-CN-S": {
        "description": "China -- small size item",
        "tax_rate": 0.0,
        "fees": [{"label": "Packaging/Shipping", "amount": 1.0}],
    },
    "IT-CN-S FRESCO": {
        "description": "China -- Frescoone-branded item",
        "tax_rate": 0.0,
        "fees": [
            {"label": "Packaging/Shipping", "amount": 2.5},
            {"label": "Engraving Logo", "amount": 0.0},
            {"label": "Custom Packaging", "amount": 0.0},
        ],
    },
}


def latest(pattern):
    matches = sorted(glob.glob(str(BASE_DIR / pattern)), key=lambda p: Path(p).stat().st_mtime)
    if not matches:
        raise FileNotFoundError(f"No file matching {pattern} in {BASE_DIR}")
    return matches[-1]


def clean_cols(df):
    df.columns = [re.sub(r"\s+", " ", str(c).replace("\n", " ")).strip() for c in df.columns]
    return df


def parse_pct(v):
    if pd.isna(v):
        return None
    s = str(v).strip().rstrip("%")
    try:
        return float(s) / 100.0
    except ValueError:
        return None


def blank_sku_record(**overrides):
    rec = {
        "itec_code": None, "currency": None, "upc": None, "status": None,
        "import_price": None, "tax": None, "fees": None,
        "cost": None, "rsp": None, "dsp": None,
        "offline_set_margin": None, "offline_purchase_price": None,
        "brand": None, "category": None, "series": None, "variant": None, "color": None,
        "date_created": None, "date_modified": None, "notes": "",
    }
    rec.update(overrides)
    return rec


def main():
    main_csv = latest("cogs-reference-google-sheet-database_*.csv")
    pp_csv = latest("database-dsp-pp_*.csv")
    print(f"Main Database export : {Path(main_csv).name}")
    print(f"Database-DSP-PP export: {Path(pp_csv).name}")

    df = clean_cols(pd.read_csv(main_csv))
    df["SKU"] = df["SKU"].astype(str).str.strip()
    df = df[df["SKU"].notna() & (df["SKU"] != "") & (df["SKU"] != "nan")]
    # Same duplicate-SKU handling as load_sku_master(): sort by _DATE so that
    # when a SKU appears more than once (re-added/re-priced later), the dict
    # below ends up keeping the most recent row, not an arbitrary one.
    df["_DATE_sort"] = pd.to_datetime(df["_DATE"], errors="coerce")
    df = df.sort_values("_DATE_sort", na_position="first")

    skus = {}
    legacy_no_itec = 0
    for _, row in df.iterrows():
        sku = row["SKU"]
        itec_code = row.get("ITEC _CODE")
        itec_code = None if pd.isna(itec_code) else str(itec_code).strip()
        date_val = row.get("_DATE")
        date_str = None
        if not pd.isna(date_val):
            try:
                date_str = pd.to_datetime(date_val).strftime("%Y-%m-%d")
            except (ValueError, TypeError):
                date_str = str(date_val)

        # Cost/DSP are stored FROZEN, exactly like the sheet's "Add To DB" button
        # pastes values not formulas -- this is deliberate, not a shortcut: the
        # currency rate and ITEC fees drift over time, so re-deriving Cost from
        # *today's* tables would silently re-price old stock at a rate/fee that
        # was never actually paid. Tax/currency actually used are carried along
        # for display/audit only; itec_code is kept as-is even if it's no longer
        # one of the 3 active China presets (IT-MLY/IT-TW/IT-CN(US)/IT-KOR/
        # IT-CN-SS/IT-CN-S FRESCO CARD), so historical records stay traceable.
        if itec_code is None or itec_code not in OLD_CODE_CURRENCY:
            legacy_no_itec += 1

        skus[sku] = blank_sku_record(
            itec_code=itec_code,
            currency=OLD_CODE_CURRENCY.get(itec_code),
            upc=None if pd.isna(row.get("UPC")) else str(row.get("UPC")).strip(),
            status=None if pd.isna(row.get("Status")) else str(row.get("Status")).strip(),
            import_price=None if pd.isna(row.get("Import Price")) else float(row.get("Import Price")),
            tax=None if pd.isna(row.get("Tax")) else float(row.get("Tax")),
            cost=None if pd.isna(row.get("Cost")) else float(row.get("Cost")),
            rsp=None if pd.isna(row.get("RSP")) else float(row.get("RSP")),
            dsp=None if pd.isna(row.get("DSP")) else float(row.get("DSP")),
            date_created=date_str, date_modified=date_str,
        )

    print(f"Migrated {len(skus)} SKUs from the main Database tab ({legacy_no_itec} without a usable/active ITEC_CODE -- frozen cost/dsp still carried over fine).")

    # Fold in the offline/physical-partner Set Margin per SKU.
    ppdf = clean_cols(pd.read_csv(pp_csv))
    ppdf["SKU"] = ppdf["SKU"].astype(str).str.strip()
    matched_pp = 0
    for _, row in ppdf.iterrows():
        sku = row["SKU"]
        set_margin = parse_pct(row.get("Set Margin"))
        purchase_price = None if pd.isna(row.get("Purchase Price")) else float(row.get("Purchase Price"))
        if set_margin is None and purchase_price is None:
            continue
        if sku in skus:
            skus[sku]["offline_set_margin"] = set_margin
            skus[sku]["offline_purchase_price"] = purchase_price
            matched_pp += 1
        else:
            # SKU only exists in the offline pricing tab -- keep it, minimal record.
            skus[sku] = blank_sku_record(
                import_price=None if pd.isna(row.get("IMPORT Price")) else float(row.get("IMPORT Price")),
                cost=None if pd.isna(row.get("Cost")) else float(row.get("Cost")),
                rsp=None if pd.isna(row.get("RSP")) else float(row.get("RSP")),
                dsp=None if pd.isna(row.get("DSP")) else float(row.get("DSP")),
                offline_set_margin=set_margin, offline_purchase_price=purchase_price,
                notes="offline-only, from Database-DSP-PP",
            )
    print(f"Matched offline Set Margin/Purchase Price for {matched_pp} SKUs ({len(ppdf) - matched_pp} in the PP export were new/unmatched).")

    # Fold in existing cost_overrides.json (manual patches from the dashboard's
    # Pending Action table) -- these win over the migrated Cost, same precedence
    # as before.
    overrides_path = BASE_DIR / "cost_overrides.json"
    folded = 0
    if overrides_path.exists():
        overrides = json.loads(overrides_path.read_text(encoding="utf-8"))
        for sku, cost in overrides.items():
            sku = str(sku).strip()
            try:
                cost = float(cost)
            except (TypeError, ValueError):
                continue
            if sku not in skus:
                skus[sku] = blank_sku_record(status="MANUAL OVERRIDE", notes="added via cost_overrides.json")
            skus[sku]["cost"] = cost
            folded += 1
        print(f"Folded in {folded} manual cost override(s) from cost_overrides.json.")

    out = {"schema_version": 2, "currencies": NEW_CURRENCIES, "itec_codes": NEW_ITEC_CODES, "skus": skus}
    out_path = BASE_DIR / "isku_database.json"
    out_path.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print(f"Wrote {out_path} (schema v2) -- {len(skus)} SKUs, {len(NEW_ITEC_CODES)} active ITEC codes, {len(NEW_CURRENCIES)} currencies.")


if __name__ == "__main__":
    main()
