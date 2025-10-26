# compare_reports_static_fixed_v2.py
"""
Fixed version:
- Normalizes mapping/DCM/Celtra keys (removes commas, trims) so IDs match across sheets.
- When selecting Celtra rows for a group, uses the mapped Celtra ID (not the group_key if group_key was a DCM id).
- Ensures __dcm_key and __celtra_key are normalized in the aggregated frames.
- Safely builds union of dates and fills missing numeric values with 0.
- Always writes a Summary sheet to avoid Excel writer errors when no sheets exist.
Usage:
    python compare_reports_static_fixed_v2.py /path/to/input.xlsx
"""

import sys
from pathlib import Path
import pandas as pd
import numpy as np

# ------------------ static column names (leave input files unchanged) ------------------
# Mapping
M_BLIS = "Blis Creative ID"
M_CELTRA = "Celtra Placement ID"
M_DCM = "DCM Placement ID"

# Blis sheet
B_DATE = "Date"
B_CREATIVE = "Blis Creative ID"
B_REQ = "Blis Requested Impression"
B_SHOWN = "Blis Shown Impression"
B_CLICKS = "Blis clicks"
B_RAW = "Blis Raw clicks"
B_WIN = "Blis win count"

# DCM sheet
D_DATE = "Date"
D_PLACEMENT = "DCM Placement ID"
D_IMP = "DCM impression"
D_CLICK = "DCM click"
D_INVALID = "DCM invalid Click"  # optional

# Celtra sheet
C_DATE = "Date"
C_PLACEMENT = "Celtra Placement ID"
C_REQ = "Celtra Requested Impression"
C_LOADED = "Celtra Loaded Impression"
C_RENDERED = "Celtra Rendered Impression"
C_CLICKS = "Clicks"  # becomes "Celtra Clicks" in output

# ------------------ helpers ------------------
def to_ymd(series):
    return pd.to_datetime(series, errors='coerce').dt.strftime("%Y-%m-%d")

def pct_safe_num(numer_arr, denom_arr):
    d = denom_arr
    b = numer_arr
    return np.where(d == 0, np.nan, (b - d) / d * 100)

def safe_sum(series):
    return 0 if (series is None or len(series) == 0) else series.sum()

def normalize_id(x):
    if pd.isna(x):
        return ''
    return str(x).replace(',', '').strip()

# ------------------ main ------------------
def make_comparison(input_xlsx: str):
    p = Path(input_xlsx)
    assert p.exists(), f"File not found: {p}"
    out_p = p.with_name(p.stem + "_comparison_static_fixed_v2.xlsx")

    xl = pd.ExcelFile(p)
    sheets = {name.lower(): name for name in xl.sheet_names}

    def read_sheet(name):
        return pd.read_excel(xl, sheet_name=sheets[name.lower()]) if name.lower() in sheets else None

    # read sheets
    mapping = read_sheet("Mapping")
    blis = read_sheet("Blis")
    dcm = read_sheet("DCM")
    celtra = read_sheet("Celtra")

    if mapping is None:
        raise ValueError("Mapping sheet (Mapping) is required.")

    # normalize mapping and ensure required mapping columns exist
    for col in [M_BLIS, M_CELTRA, M_DCM]:
        if col not in mapping.columns:
            raise ValueError(f"Mapping must contain column: '{col}' (even if some values are empty).")

    # normalize mapping ids (remove commas, trim)
    mapping[M_BLIS] = mapping[M_BLIS].astype(str).apply(lambda x: normalize_id(x))
    mapping[M_CELTRA] = mapping[M_CELTRA].astype(str).apply(lambda x: normalize_id(x))
    mapping[M_DCM] = mapping[M_DCM].astype(str).apply(lambda x: normalize_id(x))

    # ------------------ prepare Blis aggregation (creative + date) ------------------
    if blis is not None:
        if B_CREATIVE not in blis.columns:
            raise ValueError(f"Blis sheet must have '{B_CREATIVE}' column.")
        blis['__blis_key'] = blis[B_CREATIVE].astype(str).apply(lambda x: normalize_id(x))
        blis['date'] = to_ymd(blis[B_DATE]) if B_DATE in blis.columns else to_ymd(blis.iloc[:,0])
        # ensure numeric columns exist
        for c in [B_REQ, B_SHOWN, B_CLICKS, B_RAW, B_WIN]:
            if c not in blis.columns:
                blis[c] = 0
            blis[c] = pd.to_numeric(blis[c], errors='coerce').fillna(0)
        blis_agg = blis.groupby(['__blis_key','date'], as_index=False).agg({
            B_REQ: 'sum', B_SHOWN: 'sum', B_CLICKS: 'sum', B_RAW: 'sum', B_WIN: 'sum'
        }).rename(columns={
            B_REQ: "Blis Requested Impression",
            B_SHOWN: "Blis Shown Impression",
            B_CLICKS: "Blis clicks",
            B_RAW: "Blis Raw clicks",
            B_WIN: "Blis win count"
        })
    else:
        blis_agg = pd.DataFrame(columns=['__blis_key','date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])

    # ------------------ prepare DCM aggregation (placement + date) ------------------
    if dcm is not None:
        if D_PLACEMENT not in dcm.columns:
            raise ValueError(f"DCM sheet must have '{D_PLACEMENT}' column.")
        # normalize keys in DCM source as well
        dcm['__dcm_key'] = dcm[D_PLACEMENT].astype(str).apply(lambda x: normalize_id(x))
        dcm['date'] = to_ymd(dcm[D_DATE]) if D_DATE in dcm.columns else to_ymd(dcm.iloc[:,0])
        for c in [D_IMP, D_CLICK, D_INVALID]:
            if c not in dcm.columns:
                dcm[c] = 0
            dcm[c] = pd.to_numeric(dcm[c], errors='coerce').fillna(0)
        dcm_agg = dcm.groupby(['__dcm_key','date'], as_index=False).agg({
            D_IMP: 'sum', D_CLICK: 'sum', D_INVALID: 'sum'
        }).rename(columns={D_IMP: "DCM impression", D_CLICK: "DCM click", D_INVALID: "DCM invalid Click"})
        # normalize aggregated key column as string (already normalized above, but ensure type)
        dcm_agg['__dcm_key'] = dcm_agg['__dcm_key'].astype(str).apply(lambda x: normalize_id(x))
    else:
        dcm_agg = pd.DataFrame(columns=['__dcm_key','date',"DCM impression","DCM click","DCM invalid Click"])

    # ------------------ prepare Celtra aggregation (placement + date) ------------------
    if celtra is not None:
        if C_PLACEMENT not in celtra.columns:
            raise ValueError(f"Celtra sheet must have '{C_PLACEMENT}' column.")
        # normalize celtra placement ids
        celtra['__celtra_key'] = celtra[C_PLACEMENT].astype(str).apply(lambda x: normalize_id(x))
        celtra['date'] = to_ymd(celtra[C_DATE]) if C_DATE in celtra.columns else to_ymd(celtra.iloc[:,0])
        # numeric columns (create if missing)
        for src_col in [C_REQ, C_LOADED, C_RENDERED]:
            if src_col not in celtra.columns:
                celtra[src_col] = 0
            celtra[src_col] = pd.to_numeric(celtra[src_col], errors='coerce').fillna(0)
        # clicks
        celtra['Celtra Clicks'] = pd.to_numeric(celtra[C_CLICKS], errors='coerce').fillna(0) if C_CLICKS in celtra.columns else 0
        celtra_agg = celtra.groupby(['__celtra_key','date'], as_index=False).agg({
            C_REQ: 'sum', C_LOADED: 'sum', C_RENDERED: 'sum', 'Celtra Clicks': 'sum'
        }).rename(columns={C_REQ: "Celtra Requested Impression", C_LOADED: "Celtra loaded impression", C_RENDERED: "celtra Rendered Impression"})
        celtra_agg['__celtra_key'] = celtra_agg['__celtra_key'].astype(str).apply(lambda x: normalize_id(x))
    else:
        celtra_agg = pd.DataFrame(columns=['__celtra_key','date',"Celtra Requested Impression","Celtra loaded impression","celtra Rendered Impression","Celtra Clicks"])

    # ------------------ build groups from mapping (mapping ONLY links Blis -> Celtra -> DCM) ------------------
    # group_key preference: use DCM placement id if present, else Celtra placement id, else Blis creative id
    def choose_group_key(dcm_val, cel_val, blis_val):
        d = normalize_id(dcm_val)
        c = normalize_id(cel_val)
        b = normalize_id(blis_val)
        if d and d.lower() != 'nan':
            return d
        if c and c.lower() != 'nan':
            return c
        return b

    mapping['group_key'] = mapping.apply(lambda r: choose_group_key(r[M_DCM], r[M_CELTRA], r[M_BLIS]), axis=1)
    # for each group key, store the DCM and Celtra ID (if any) for display in sheet columns (normalized)
    group_to_dcm = mapping.groupby('group_key')[M_DCM].first().to_dict()
    group_to_celtra = mapping.groupby('group_key')[M_CELTRA].first().to_dict()
    # normalize those dict values as well
    group_to_dcm = {k: normalize_id(v) for k, v in group_to_dcm.items()}
    group_to_celtra = {k: normalize_id(v) for k, v in group_to_celtra.items()}

    # map Blis creative -> group_key (normalized)
    blis_to_group = dict(zip(mapping[M_BLIS].astype(str).apply(lambda x: normalize_id(x)), mapping['group_key']))

    # ------------------ map Blis aggregated to groups and aggregate by group+date ----------------
    if not blis_agg.empty:
        blis_agg['__blis_key'] = blis_agg['__blis_key'].astype(str).apply(lambda x: normalize_id(x))
        blis_agg['group_key'] = blis_agg['__blis_key'].map(blis_to_group)
        blis_by_group = blis_agg.groupby(['group_key','date'], as_index=False).agg({
            "Blis Requested Impression": 'sum',
            "Blis Shown Impression": 'sum',
            "Blis clicks": 'sum',
            "Blis Raw clicks": 'sum',
            "Blis win count": 'sum'
        })
    else:
        blis_by_group = pd.DataFrame(columns=['group_key','date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])

    # ------------------ output workbook ------------------
    unique_group_keys = mapping['group_key'].unique().tolist()
    summary_rows = []

    with pd.ExcelWriter(out_p, engine='openpyxl') as writer:
        # optional mapping warnings: same Blis creative mapped to multiple groups?
        dup_check = mapping.groupby(M_BLIS).agg({M_CELTRA: pd.Series.nunique, M_DCM: pd.Series.nunique})
        dup_warn = dup_check[(dup_check[M_CELTRA] > 1) | (dup_check[M_DCM] > 1)]
        if not dup_warn.empty:
            dup_warn.reset_index().to_excel(writer, sheet_name='Mapping_Warnings', index=False)

        for key in unique_group_keys:
            # normalized group key
            key_norm = normalize_id(key)

            # retrieve the mapped DCM and Celtra ids for this group (normalized)
            dcm_id_for_group = group_to_dcm.get(key_norm)
            cel_id_for_group = group_to_celtra.get(key_norm)

            # Now pick dcm_frame by the DCM id (if present)
            if not dcm_agg.empty and dcm_id_for_group:
                dcm_frame = dcm_agg[dcm_agg['__dcm_key'] == dcm_id_for_group].copy()
            else:
                dcm_frame = pd.DataFrame(columns=dcm_agg.columns) if not dcm_agg.empty else pd.DataFrame()

            # Pick celtra_frame by the celtra id (if present)
            if not celtra_agg.empty and cel_id_for_group:
                cel_frame = celtra_agg[celtra_agg['__celtra_key'] == cel_id_for_group].copy()
            else:
                cel_frame = pd.DataFrame(columns=celtra_agg.columns) if not celtra_agg.empty else pd.DataFrame()

            # Blis frame is mapped by group_key already (use normalized key)
            blis_frame = blis_by_group[blis_by_group['group_key'] == key_norm].copy() if (not blis_by_group.empty and key_norm in blis_by_group['group_key'].values) else pd.DataFrame()

            # ---------------- SAFE: build union of dates ----------------
            date_sets = []
            for df in (dcm_frame, cel_frame, blis_frame):
                if isinstance(df, pd.DataFrame) and 'date' in df.columns and not df['date'].dropna().empty:
                    # cast to str to avoid dtype issues
                    date_sets.append(set(df['date'].dropna().astype(str).tolist()))
                else:
                    date_sets.append(set())

            all_dates = sorted(set().union(*date_sets))

            # if no dates at all, skip this group but add empty summary row
            if not all_dates:
                summary_rows.append({
                    'group_key': key_norm,
                    'dcm_present': not dcm_frame.empty,
                    'celtra_present': not cel_frame.empty,
                    'blis_present': not blis_frame.empty,
                    'dcm_impressions_total': 0,
                    'blis_impressions_total': 0,
                    'celtra_rendered_total': 0,
                    'blis_vs_dcm_pct_total': np.nan,
                    'blis_vs_celtra_pct_total': np.nan
                })
                continue

            # build dates dataframe (ensure date strings are consistent)
            dates_df = pd.DataFrame({'date': all_dates})

            # Merge aggregated frames (only if they contain a 'date' column)
            merged = dates_df.copy()

            if isinstance(dcm_frame, pd.DataFrame) and 'date' in dcm_frame.columns and not dcm_frame.empty:
                merged = merged.merge(dcm_frame[['date', "DCM impression", "DCM click", "DCM invalid Click"]], on='date', how='left')
            else:
                merged[["DCM impression", "DCM click", "DCM invalid Click"]] = 0

            if isinstance(blis_frame, pd.DataFrame) and 'date' in blis_frame.columns and not blis_frame.empty:
                merged = merged.merge(blis_frame[['date', "Blis Requested Impression", "Blis Shown Impression", "Blis clicks", "Blis Raw clicks", "Blis win count"]], on='date', how='left')
            else:
                merged[["Blis Requested Impression", "Blis Shown Impression", "Blis clicks", "Blis Raw clicks", "Blis win count"]] = 0

            if isinstance(cel_frame, pd.DataFrame) and 'date' in cel_frame.columns and not cel_frame.empty:
                merged = merged.merge(cel_frame[['date', "Celtra Requested Impression", "Celtra loaded impression", "celtra Rendered Impression", "Celtra Clicks"]], on='date', how='left')
            else:
                merged[["Celtra Requested Impression", "Celtra loaded impression", "celtra Rendered Impression", "Celtra Clicks"]] = 0

            # normalize numeric columns to numeric and fill zeros
            out_numeric = ["DCM impression", "DCM click", "DCM invalid Click",
                           "Blis Requested Impression", "Blis Shown Impression", "Blis clicks", "Blis Raw clicks", "Blis win count",
                           "Celtra Requested Impression", "Celtra loaded impression", "celtra Rendered Impression", "Celtra Clicks"]
            for c in out_numeric:
                if c in merged.columns:
                    merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0)
                else:
                    merged[c] = 0

            # explicitly add the placement ID columns (show the actual mapped IDs)
            dcm_id_display = dcm_id_for_group if dcm_id_for_group else np.nan
            cel_id_display = cel_id_for_group if cel_id_for_group else np.nan
            merged["DCM placementID"] = dcm_id_display
            merged["Celtra placementID"] = cel_id_display

            # --- percent/diff calculations (only if relevant columns exist) ---
            # blis vs DCM
            merged['blis vs DCM impression %'] = pct_safe_num(merged['Blis Requested Impression'].values, merged['DCM impression'].values)
            merged['blis vs DCM click %'] = pct_safe_num(merged['Blis clicks'].values, merged['DCM click'].values)
            # raw vs invalid
            if 'DCM invalid Click' in merged.columns:
                merged['Blis raw click vs DCM invalid %'] = pct_safe_num(merged['Blis Raw clicks'].values, merged['DCM invalid Click'].values)
            # blis vs celtra loaded
            merged['blis impression vs celtra loaded %'] = pct_safe_num(merged['Blis Requested Impression'].values, merged['Celtra loaded impression'].values) if 'Celtra loaded impression' in merged.columns else np.nan
            merged['blis click vs Celtra click %'] = pct_safe_num(merged['Blis clicks'].values, merged['Celtra Clicks'].values) if 'Celtra Clicks' in merged.columns else np.nan
            # celtra loaded vs DCM impression
            merged['celtra loaded vs DCM impression %'] = pct_safe_num(merged['Celtra loaded impression'].values, merged['DCM impression'].values) if 'Celtra loaded impression' in merged.columns else np.nan
            # celtra clicks vs DCM clicks
            merged['Celtra click vs DCM click %'] = pct_safe_num(merged['Celtra Clicks'].values, merged['DCM click'].values) if 'Celtra Clicks' in merged.columns else np.nan

            # reorder exactly per your requested mandatory columns, include only existing cols
            desired_order = [
                'date', 'DCM placementID', 'Celtra placementID',
                'Blis Requested Impression','Blis Shown Impression','Blis clicks','Blis Raw clicks','Blis win count',
                'DCM impression','DCM click','DCM invalid Click',
                'Celtra Requested Impression','Celtra loaded impression','celtra Rendered Impression','Celtra Clicks',
                'blis vs DCM impression %','blis vs DCM click %','Blis raw click vs DCM invalid %',
                'blis impression vs celtra loaded %','blis click vs Celtra click %','celtra loaded vs DCM impression %','Celtra click vs DCM click %'
            ]
            final_cols = [c for c in desired_order if c in merged.columns]
            merged = merged[final_cols]

            # round percent columns for readability
            for c in merged.columns:
                if '%' in c:
                    merged[c] = merged[c].round(2)

            # ---------------- Grand Total ----------------
            numeric_for_sum = [c for c in merged.columns if c not in ('date','DCM placementID','Celtra placementID') and merged[c].dtype.kind in 'biufc']
            totals = {}
            for c in numeric_for_sum:
                totals[c] = safe_sum(merged[c])

            # compute derived percent totals (safe)
            def pct_total(numer_total, denom_total):
                return round((numer_total - denom_total) / denom_total * 100, 2) if denom_total != 0 else np.nan

            if 'Blis Requested Impression' in merged.columns and 'DCM impression' in merged.columns:
                totals['blis vs DCM impression %'] = pct_total(totals.get('Blis Requested Impression',0), totals.get('DCM impression',0))
            if 'Blis clicks' in merged.columns and 'DCM click' in merged.columns:
                totals['blis vs DCM click %'] = pct_total(totals.get('Blis clicks',0), totals.get('DCM click',0))
            if 'Blis Raw clicks' in merged.columns and 'DCM invalid Click' in merged.columns:
                totals['Blis raw click vs DCM invalid %'] = pct_total(totals.get('Blis Raw clicks',0), totals.get('DCM invalid Click',0))
            if 'Blis Requested Impression' in merged.columns and 'Celtra loaded impression' in merged.columns:
                totals['blis impression vs celtra loaded %'] = pct_total(totals.get('Blis Requested Impression',0), totals.get('Celtra loaded impression',0))
            if 'Blis clicks' in merged.columns and 'Celtra Clicks' in merged.columns:
                totals['blis click vs Celtra click %'] = pct_total(totals.get('Blis clicks',0), totals.get('Celtra Clicks',0))
            if 'Celtra loaded impression' in merged.columns and 'DCM impression' in merged.columns:
                totals['celtra loaded vs DCM impression %'] = pct_total(totals.get('Celtra loaded impression',0), totals.get('DCM impression',0))
            if 'Celtra Clicks' in merged.columns and 'DCM click' in merged.columns:
                totals['Celtra click vs DCM click %'] = pct_total(totals.get('Celtra Clicks',0), totals.get('DCM click',0))

            # Placement ID values in totals row
            totals['DCM placementID'] = dcm_id_display if not pd.isna(dcm_id_display) else np.nan
            totals['Celtra placementID'] = cel_id_display if not pd.isna(cel_id_display) else np.nan
            totals['date'] = "Grand Total"

            total_row = pd.DataFrame([totals])
            total_row = total_row.reindex(columns=merged.columns, fill_value=0)
            # keep percent NaNs; round percent cols
            for c in merged.columns:
                if '%' in c and c in total_row.columns:
                    v = total_row.at[0,c]
                    total_row.at[0,c] = round(v,2) if pd.notna(v) else v

            merged = pd.concat([merged, total_row], ignore_index=True, sort=False)

            # write sheet
            sheet_name = f"Group_{key_norm}"[:31]
            merged.to_excel(writer, sheet_name=sheet_name, index=False)

            # summary row using computed totals
            summary_rows.append({
                'group_key': key_norm,
                'dcm_present': not dcm_frame.empty,
                'celtra_present': not cel_frame.empty,
                'blis_present': not blis_frame.empty,
                'dcm_impressions_total': totals.get('DCM impression', 0),
                'blis_impressions_total': totals.get('Blis Requested Impression', 0),
                'Celtra loaded impression': totals.get('Celtra loaded impression', 0),
                'blis_vs_dcm_pct_total': totals.get('blis vs DCM impression %'),
                'blis_vs_celtra_pct_total': totals.get('blis impression vs celtra loaded %')
            })

        # write summary (always write to avoid empty-workbook errors)
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False)

    print("Done. Output written to:", out_p)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python compare_reports_static_fixed_v2.py /path/to/input.xlsx")
        sys.exit(1)
    make_comparison(sys.argv[1])
