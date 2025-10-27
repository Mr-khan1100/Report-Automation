# compare_reports_static_fixed_v3.py
"""
Handles mappings where:
 - 1 DCM placement -> many Celtra placements
 - 1 Celtra placement -> many Blis creatives
 - Each Blis creative maps to at most one Celtra and one DCM

Usage:
    python compare_reports_static_fixed_v3.py /path/to/input.xlsx
"""
import sys
from pathlib import Path
import pandas as pd
import numpy as np

# ------------------ static column names ------------------
M_BLIS = "Blis Creative ID"
M_CELTRA = "Celtra Placement ID"
M_DCM = "DCM Placement ID"

B_DATE = "Date"
B_CREATIVE = "Blis Creative ID"
B_REQ = "Blis Requested Impression"
B_SHOWN = "Blis Shown Impression"
B_CLICKS = "Blis clicks"
B_RAW = "Blis Raw clicks"
B_WIN = "Blis win count"

D_DATE = "Date"
D_PLACEMENT = "DCM Placement ID"
D_IMP = "DCM impression"
D_CLICK = "DCM click"
D_INVALID = "DCM invalid Click"  # optional

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
    out_p = p.with_name(p.stem + "_comparison_static_fixed_v3.xlsx")

    xl = pd.ExcelFile(p)
    sheets = {name.lower(): name for name in xl.sheet_names}
    def read_sheet(name):
        return pd.read_excel(xl, sheet_name=sheets[name.lower()]) if name.lower() in sheets else None

    mapping = read_sheet("Mapping")
    blis = read_sheet("Blis")
    dcm = read_sheet("DCM")
    celtra = read_sheet("Celtra")

    if mapping is None:
        raise ValueError("Mapping sheet (Mapping) is required.")

    # ensure mapping columns present
    for col in [M_BLIS, M_CELTRA, M_DCM]:
        if col not in mapping.columns:
            raise ValueError(f"Mapping must contain column: '{col}' (even if some values are empty).")

    # normalize mapping ids
    mapping[M_BLIS] = mapping[M_BLIS].astype(str).apply(normalize_id)
    mapping[M_CELTRA] = mapping[M_CELTRA].astype(str).apply(normalize_id)
    mapping[M_DCM] = mapping[M_DCM].astype(str).apply(normalize_id)

    # ------------------ prepare Blis aggregation (creative + date) ------------------
    if blis is not None:
        if B_CREATIVE not in blis.columns:
            raise ValueError(f"Blis sheet must have '{B_CREATIVE}' column.")
        blis['__blis_key'] = blis[B_CREATIVE].astype(str).apply(normalize_id)
        blis['date'] = to_ymd(blis[B_DATE]) if B_DATE in blis.columns else to_ymd(blis.iloc[:,0])
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
        blis_agg['__blis_key'] = blis_agg['__blis_key'].astype(str).apply(normalize_id)
    else:
        blis_agg = pd.DataFrame(columns=['__blis_key','date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])

    # ------------------ prepare DCM aggregation (placement + date) ------------------
    if dcm is not None:
        if D_PLACEMENT not in dcm.columns:
            raise ValueError(f"DCM sheet must have '{D_PLACEMENT}' column.")
        dcm['__dcm_key'] = dcm[D_PLACEMENT].astype(str).apply(normalize_id)
        dcm['date'] = to_ymd(dcm[D_DATE]) if D_DATE in dcm.columns else to_ymd(dcm.iloc[:,0])
        for c in [D_IMP, D_CLICK, D_INVALID]:
            if c not in dcm.columns:
                dcm[c] = 0
            dcm[c] = pd.to_numeric(dcm[c], errors='coerce').fillna(0)
        dcm_agg = dcm.groupby(['__dcm_key','date'], as_index=False).agg({
            D_IMP: 'sum', D_CLICK: 'sum', D_INVALID: 'sum'
        }).rename(columns={D_IMP: "DCM impression", D_CLICK: "DCM click", D_INVALID: "DCM invalid Click"})
        dcm_agg['__dcm_key'] = dcm_agg['__dcm_key'].astype(str).apply(normalize_id)
    else:
        dcm_agg = pd.DataFrame(columns=['__dcm_key','date',"DCM impression","DCM click","DCM invalid Click"])

    # ------------------ prepare Celtra aggregation (placement + date) ------------------
    if celtra is not None:
        if C_PLACEMENT not in celtra.columns:
            raise ValueError(f"Celtra sheet must have '{C_PLACEMENT}' column.")
        celtra['__celtra_key'] = celtra[C_PLACEMENT].astype(str).apply(normalize_id)
        celtra['date'] = to_ymd(celtra[C_DATE]) if C_DATE in celtra.columns else to_ymd(celtra.iloc[:,0])
        for src_col in [C_REQ, C_LOADED, C_RENDERED]:
            if src_col not in celtra.columns:
                celtra[src_col] = 0
            celtra[src_col] = pd.to_numeric(celtra[src_col], errors='coerce').fillna(0)
        celtra['Celtra Clicks'] = pd.to_numeric(celtra[C_CLICKS], errors='coerce').fillna(0) if C_CLICKS in celtra.columns else 0
        celtra_agg = celtra.groupby(['__celtra_key','date'], as_index=False).agg({
            C_REQ: 'sum', C_LOADED: 'sum', C_RENDERED: 'sum', 'Celtra Clicks': 'sum'
        }).rename(columns={C_REQ: "Celtra Requested Impression", C_LOADED: "Celtra Loaded Impression", C_RENDERED: "Celtra Rendered Impression"})
        celtra_agg['__celtra_key'] = celtra_agg['__celtra_key'].astype(str).apply(normalize_id)
    else:
        celtra_agg = pd.DataFrame(columns=['__celtra_key','date',"Celtra Requested Impression","Celtra Loaded Impression","Celtra Rendered Impression","Celtra Clicks"])

    # ------------------ build mapping lookups ------------------
    # dcm -> list of celtra ids, dcm -> list of blis creatives
    dcm_to_celtras = mapping.groupby(M_DCM)[M_CELTRA].apply(lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower()!='nan']))).to_dict()
    dcm_to_blis = mapping.groupby(M_DCM)[M_BLIS].apply(lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower()!='nan']))).to_dict()
    # celtra -> list of blis creatives
    celtra_to_blis = mapping.groupby(M_CELTRA)[M_BLIS].apply(lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower()!='nan']))).to_dict()
    # blis -> (celtra,dcm) single mapping
    blis_to_celtra = mapping.drop_duplicates(subset=[M_BLIS]).set_index(M_BLIS)[M_CELTRA].to_dict()
    blis_to_dcm = mapping.drop_duplicates(subset=[M_BLIS]).set_index(M_BLIS)[M_DCM].to_dict()

    # group keys: choose DCM if present else Celtra else Blis creative as before
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
    unique_group_keys = sorted(set(mapping['group_key'].tolist()))

    # ------------------ main loop: for each mapping-derived group produce sheet ------------------
    summary_rows = []
    with pd.ExcelWriter(out_p, engine='openpyxl') as writer:
        # mapping warnings
        dup_check = mapping.groupby(M_BLIS).agg({M_CELTRA: pd.Series.nunique, M_DCM: pd.Series.nunique})
        dup_warn = dup_check[(dup_check[M_CELTRA] > 1) | (dup_check[M_DCM] > 1)]
        if not dup_warn.empty:
            dup_warn.reset_index().to_excel(writer, sheet_name='Mapping_Warnings', index=False)

        for key in unique_group_keys:
            key_norm = normalize_id(key)

            # determine if this group is a DCM-group, Celtra-group, or Blis-group
            is_dcm_group = key_norm in mapping[M_DCM].values
            is_celtra_group = key_norm in mapping[M_CELTRA].values and not is_dcm_group
            # is_blis_group = key_norm in mapping[M_BLIS].values and not (is_dcm_group or is_celtra_group)

            # collect lists according to group type
            if is_dcm_group:
                dcm_id_for_group = key_norm
                celtra_ids = dcm_to_celtras.get(dcm_id_for_group, [])
                blis_list = dcm_to_blis.get(dcm_id_for_group, [])
            elif is_celtra_group:
                cel_id_for_group = key_norm
                dcm_vals = mapping.loc[mapping[M_CELTRA] == cel_id_for_group, M_DCM].dropna().unique().tolist()
                dcm_id_for_group = normalize_id(dcm_vals[0]) if dcm_vals else ''
                celtra_ids = [cel_id_for_group]
                blis_list = celtra_to_blis.get(cel_id_for_group, [])
            else:  # blis-group
                blis_creative = key_norm
                blis_list = [blis_creative]
                cel_id = normalize_id(blis_to_celtra.get(blis_creative, ''))
                dcm_id_for_group = normalize_id(blis_to_dcm.get(blis_creative, ''))
                celtra_ids = [cel_id] if cel_id else []

            # Build dcm_frame (if dcm present)
            if not dcm_agg.empty and dcm_id_for_group:
                dcm_frame = dcm_agg[dcm_agg['__dcm_key'] == dcm_id_for_group].copy()
            else:
                dcm_frame = pd.DataFrame(columns=dcm_agg.columns) if not dcm_agg.empty else pd.DataFrame()

            # ---------- NEW: If DCM-group, create one merged per Celtra (don't collapse celtra IDs) ----------
            merged = pd.DataFrame()  # final frame for this group

            if is_dcm_group and celtra_ids:
                merged_rows = []
                # helper that builds a per-celtra merged frame (date-level), returns DataFrame or None
                def build_one_merged(cel_id, blis_list_for_cel):
                    # cel_frame for this single celtra (already aggregated)
                    if not celtra_agg.empty and cel_id:
                        cel_frame_local = celtra_agg[celtra_agg['__celtra_key'] == cel_id].copy()
                    else:
                        cel_frame_local = pd.DataFrame(columns=celtra_agg.columns) if not celtra_agg.empty else pd.DataFrame()

                    # blis_frame for this celtra: aggregate only creatives that belong to this celtra
                    if not blis_agg.empty and blis_list_for_cel:
                        bf = blis_agg[blis_agg['__blis_key'].isin(blis_list_for_cel)].copy()
                        if not bf.empty:
                            blis_frame_local = bf.groupby('date', as_index=False).agg({
                                "Blis Requested Impression": 'sum',
                                "Blis Shown Impression": 'sum',
                                "Blis clicks": 'sum',
                                "Blis Raw clicks": 'sum',
                                "Blis win count": 'sum'
                            })
                        else:
                            blis_frame_local = pd.DataFrame(columns=['date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])
                    else:
                        blis_frame_local = pd.DataFrame(columns=['date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])

                    # Build union dates (DCM dates + celtra_local + blis_local)
                    date_sets_local = []
                    for df_local in (dcm_frame, cel_frame_local, blis_frame_local):
                        if isinstance(df_local, pd.DataFrame) and 'date' in df_local.columns and not df_local['date'].dropna().empty:
                            date_sets_local.append(set(df_local['date'].dropna().astype(str).tolist()))
                        else:
                            date_sets_local.append(set())
                    all_dates_local = sorted(set().union(*date_sets_local))
                    if not all_dates_local:
                        return None  # nothing to add for this celtra

                    dates_df_local = pd.DataFrame({'date': all_dates_local})
                    merged_local = dates_df_local.copy()

                    # Merge date-level DCM (same dcm_frame for all celtras)
                    if not dcm_frame.empty:
                        merged_local = merged_local.merge(dcm_frame[['date', "DCM impression", "DCM click", "DCM invalid Click"]], on='date', how='left')
                    else:
                        merged_local[["DCM impression", "DCM click", "DCM invalid Click"]] = 0

                    # Merge blis specific to this celtra
                    if not blis_frame_local.empty:
                        merged_local = merged_local.merge(blis_frame_local[['date', "Blis Requested Impression", "Blis Shown Impression", "Blis clicks", "Blis Raw clicks", "Blis win count"]], on='date', how='left')
                    else:
                        merged_local[["Blis Requested Impression", "Blis Shown Impression", "Blis clicks", "Blis Raw clicks", "Blis win count"]] = 0

                    # Merge celtra (single celtra row)
                    if not cel_frame_local.empty:
                        merged_local = merged_local.merge(cel_frame_local[['date', "Celtra Requested Impression", "Celtra Loaded Impression", "Celtra Rendered Impression", "Celtra Clicks"]], on='date', how='left')
                    else:
                        merged_local[["Celtra Requested Impression", "Celtra Loaded Impression", "Celtra Rendered Impression", "Celtra Clicks"]] = 0

                    # add placement columns (DCM + this celtra)
                    merged_local["DCM placementID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                    merged_local["Celtra placementID"] = cel_id if cel_id else np.nan

                    return merged_local

                # build rows per celtra id
                for cel_id in celtra_ids:
                    blis_list_for_cel = celtra_to_blis.get(cel_id, [])
                    ml = build_one_merged(cel_id, blis_list_for_cel)
                    if ml is not None:
                        merged_rows.append(ml)

                # If nothing produced (no dates), fallback to a DCM-only merged frame (single row per date)
                if not merged_rows:
                    # Build a DCM-only merged (no celtra, no blis)
                    # union dates with dcm_frame only
                    if not dcm_frame.empty and 'date' in dcm_frame.columns and not dcm_frame['date'].dropna().empty:
                        dates_df_local = pd.DataFrame({'date': sorted(dcm_frame['date'].dropna().astype(str).tolist())})
                        merged_local = dates_df_local.merge(dcm_frame[['date', "DCM impression", "DCM click", "DCM invalid Click"]], on='date', how='left')
                        merged_local[["Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"]] = 0
                        merged_local[["Celtra Requested Impression","Celtra Loaded Impression","Celtra Rendered Impression","Celtra Clicks"]] = 0
                        merged_local["DCM placementID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                        merged_local["Celtra placementID"] = np.nan
                        merged_rows.append(merged_local)

                # concat per-celtra frames
                merged = pd.concat(merged_rows, ignore_index=True, sort=False) if merged_rows else pd.DataFrame()

            else:
                # Non-DCM group (Celtra-group or Blis-group) keep previous aggregated single-row-per-date behavior
                # Build cel_frame combined (single celtra) and blis_frame combined (single set) then merge once
                if celtra_ids:
                    cel_frame = celtra_agg[celtra_agg['__celtra_key'].isin(celtra_ids)].copy()
                    if not cel_frame.empty:
                        cel_frame = cel_frame.groupby('date', as_index=False).agg({
                            "Celtra Requested Impression": 'sum',
                            "Celtra Loaded Impression": 'sum',
                            "Celtra Rendered Impression": 'sum',
                            "Celtra Clicks": 'sum'
                        })
                else:
                    cel_frame = pd.DataFrame(columns=celtra_agg.columns) if not celtra_agg.empty else pd.DataFrame()

                if not blis_agg.empty and blis_list:
                    bf = blis_agg[blis_agg['__blis_key'].isin(blis_list)].copy()
                    if not bf.empty:
                        blis_frame = bf.groupby('date', as_index=False).agg({
                            "Blis Requested Impression": 'sum',
                            "Blis Shown Impression": 'sum',
                            "Blis clicks": 'sum',
                            "Blis Raw clicks": 'sum',
                            "Blis win count": 'sum'
                        })
                    else:
                        blis_frame = pd.DataFrame(columns=['date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])
                else:
                    blis_frame = pd.DataFrame(columns=['date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"])

                # Now build the single merged as before
                date_sets = []
                for df in (dcm_frame, cel_frame, blis_frame):
                    if isinstance(df, pd.DataFrame) and 'date' in df.columns and not df['date'].dropna().empty:
                        date_sets.append(set(df['date'].dropna().astype(str).tolist()))
                    else:
                        date_sets.append(set())
                all_dates = sorted(set().union(*date_sets))
                if not all_dates:
                    merged = pd.DataFrame()
                else:
                    dates_df = pd.DataFrame({'date': all_dates})
                    merged = dates_df.copy()
                    if not dcm_frame.empty:
                        merged = merged.merge(dcm_frame[['date',"DCM impression","DCM click","DCM invalid Click"]], on='date', how='left')
                    else:
                        merged[["DCM impression","DCM click","DCM invalid Click"]] = 0
                    if not blis_frame.empty:
                        merged = merged.merge(blis_frame[['date',"Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"]], on='date', how='left')
                    else:
                        merged[["Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count"]] = 0
                    if not cel_frame.empty:
                        merged = merged.merge(cel_frame[['date',"Celtra Requested Impression","Celtra Loaded Impression","Celtra Rendered Impression","Celtra Clicks"]], on='date', how='left')
                    else:
                        merged[["Celtra Requested Impression","Celtra Loaded Impression","Celtra Rendered Impression","Celtra Clicks"]] = 0
                    # show placement ids (for single-row case)
                    merged["DCM placementID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                    merged["Celtra placementID"] = ",".join(celtra_ids) if celtra_ids else np.nan

            # If merged is empty (no data) skip writing sheet and add summary row
            if merged.empty:
                summary_rows.append({
                    'group_key': key_norm,
                    'dcm_present': not dcm_frame.empty,
                    'celtra_present': len(celtra_ids) > 0,
                    'blis_present': len(blis_list) > 0,
                    'dcm_impressions_total': 0,
                    'blis_impressions_total': 0,
                    'celtra_rendered_total': 0,
                    'blis_vs_dcm_pct_total': np.nan,
                    'blis_vs_celtra_pct_total': np.nan
                })
                continue

            # numeric coercion + zeros for merged
            out_numeric = ["DCM impression","DCM click","DCM invalid Click",
                           "Blis Requested Impression","Blis Shown Impression","Blis clicks","Blis Raw clicks","Blis win count",
                           "Celtra Requested Impression","Celtra Loaded Impression","Celtra Rendered Impression","Celtra Clicks"]
            for c in out_numeric:
                if c in merged.columns:
                    merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0)
                else:
                    merged[c] = 0

            # percent/diff calculations
            merged['blis vs DCM impression %'] = pct_safe_num(merged['Blis Requested Impression'].values, merged['DCM impression'].values)
            merged['blis vs DCM click %'] = pct_safe_num(merged['Blis clicks'].values, merged['DCM click'].values)
            if 'DCM invalid Click' in merged.columns:
                merged['Blis raw click vs DCM invalid %'] = pct_safe_num(merged['Blis Raw clicks'].values, merged['DCM invalid Click'].values)
            merged['blis impression vs celtra loaded %'] = pct_safe_num(merged['Blis Requested Impression'].values, merged['Celtra Loaded Impression'].values) if 'Celtra Loaded Impression' in merged.columns else np.nan
            merged['blis click vs Celtra click %'] = pct_safe_num(merged['Blis clicks'].values, merged['Celtra Clicks'].values) if 'Celtra Clicks' in merged.columns else np.nan
            merged['celtra loaded vs DCM impression %'] = pct_safe_num(merged['Celtra Loaded Impression'].values, merged['DCM impression'].values) if 'Celtra Loaded Impression' in merged.columns else np.nan
            merged['Celtra click vs DCM click %'] = pct_safe_num(merged['Celtra Clicks'].values, merged['DCM click'].values) if 'Celtra Clicks' in merged.columns else np.nan

            # reorder columns
            desired_order = [
                'date', 'DCM placementID', 'Celtra placementID',
                'Blis Requested Impression','Blis Shown Impression','Blis clicks','Blis Raw clicks','Blis win count',
                'DCM impression','DCM click','DCM invalid Click',
                'Celtra Requested Impression','Celtra Loaded Impression','Celtra Rendered Impression','Celtra Clicks',
                'blis vs DCM impression %','blis vs DCM click %','Blis raw click vs DCM invalid %',
                'blis impression vs celtra loaded %','blis click vs Celtra click %','celtra loaded vs DCM impression %','Celtra click vs DCM click %'
            ]
            final_cols = [c for c in desired_order if c in merged.columns]
            merged = merged[final_cols]

            # round percents
            for c in merged.columns:
                if '%' in c:
                    merged[c] = merged[c].round(2)

            # Grand total row
            numeric_for_sum = [c for c in merged.columns if c not in ('date','DCM placementID','Celtra placementID') and merged[c].dtype.kind in 'biufc']
            totals = {}
            for c in numeric_for_sum:
                totals[c] = safe_sum(merged[c])

            # If this is a DCM-group with multiple celtras, override DCM totals from dcm_frame to avoid double-counting
            if is_dcm_group and celtra_ids and not dcm_frame.empty:
                totals['DCM impression'] = dcm_frame['DCM impression'].sum() if 'DCM impression' in dcm_frame.columns else 0
                totals['DCM click'] = dcm_frame['DCM click'].sum() if 'DCM click' in dcm_frame.columns else 0
                if 'DCM invalid Click' in dcm_frame.columns:
                    totals['DCM invalid Click'] = dcm_frame['DCM invalid Click'].sum()

            def pct_total(numer_total, denom_total):
                return round((numer_total - denom_total) / denom_total * 100, 2) if denom_total != 0 else np.nan
            if 'Blis Requested Impression' in merged.columns and 'DCM impression' in merged.columns:
                totals['blis vs DCM impression %'] = pct_total(totals.get('Blis Requested Impression',0), totals.get('DCM impression',0))
            if 'Blis clicks' in merged.columns and 'DCM click' in merged.columns:
                totals['blis vs DCM click %'] = pct_total(totals.get('Blis clicks',0), totals.get('DCM click',0))
            if 'Blis Raw clicks' in merged.columns and 'DCM invalid Click' in merged.columns:
                totals['Blis raw click vs DCM invalid %'] = pct_total(totals.get('Blis Raw clicks',0), totals.get('DCM invalid Click',0))
            if 'Blis Requested Impression' in merged.columns and 'Celtra Loaded Impression' in merged.columns:
                totals['blis impression vs celtra loaded %'] = pct_total(totals.get('Blis Requested Impression',0), totals.get('Celtra Loaded Impression',0))
            if 'Blis clicks' in merged.columns and 'Celtra Clicks' in merged.columns:
                totals['blis click vs Celtra click %'] = pct_total(totals.get('Blis clicks',0), totals.get('Celtra Clicks',0))
            if 'Celtra Loaded Impression' in merged.columns and 'DCM impression' in merged.columns:
                totals['celtra loaded vs DCM impression %'] = pct_total(totals.get('Celtra Loaded Impression',0), totals.get('DCM impression',0))
            if 'Celtra Clicks' in merged.columns and 'DCM click' in merged.columns:
                totals['Celtra click vs DCM click %'] = pct_total(totals.get('Celtra Clicks',0), totals.get('DCM click',0))

            # placementID values in totals row
            totals['DCM placementID'] = dcm_id_for_group if dcm_id_for_group else np.nan
            # For DCM-group we don't want comma-joined Celtra list in totals; keep as comma list for non-DCM or empty
            totals['Celtra placementID'] = ",".join(celtra_ids) if (not is_dcm_group and celtra_ids) else (np.nan if is_dcm_group else np.nan)
            totals['date'] = "Grand Total"

            total_row = pd.DataFrame([totals])
            total_row = total_row.reindex(columns=merged.columns, fill_value=0)
            for c in merged.columns:
                if '%' in c and c in total_row.columns:
                    v = total_row.at[0,c]
                    total_row.at[0,c] = round(v,2) if pd.notna(v) else v
            merged = pd.concat([merged, total_row], ignore_index=True, sort=False)

            # write sheet
            sheet_name = f"Group_{key_norm}"[:31]
            merged.to_excel(writer, sheet_name=sheet_name, index=False)

            # summary row
            summary_rows.append({
                'group_key': key_norm,
                'dcm_present': not dcm_frame.empty,
                'celtra_present': len(celtra_ids) > 0,
                'blis_present': len(blis_list) > 0,
                'dcm_impressions_total': totals.get('DCM impression', 0),
                'blis_impressions_total': totals.get('Blis Requested Impression', 0),
                'celtra_rendered_total': totals.get('Celtra Rendered Impression', 0),
                'blis_vs_dcm_pct_total': totals.get('blis vs DCM impression %'),
                'blis_vs_celtra_pct_total': totals.get('blis impression vs celtra loaded %')
            })

        # always write summary
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False)

    print("Done. Output written to:", out_p)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python compare_reports_static_fixed_v3.py /path/to/input.xlsx")
        sys.exit(1)
    make_comparison(sys.argv[1])
