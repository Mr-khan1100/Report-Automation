# compare_reports_static_fixed_v4_dynamic_columns.py
"""
Handles mappings where:
 - 1 DCM placement -> many Celtra placements
 - 1 Celtra placement -> many Blis creatives
 - Each Blis creative maps to at most one Celtra and one DCM

Usage:
    python compare_reports_static_fixed_v4_dynamic_columns.py /path/to/input.xlsx
"""
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl.utils import get_column_letter

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
D_INVALID_IMP = "DCM invalid Impression"  # optional
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
    return pd.to_datetime(series, errors="coerce").dt.strftime("%Y-%m-%d")

def discrepancy_pct(numer_arr, denom_arr):
    numer = np.asarray(numer_arr, dtype="float64")
    denom = np.asarray(denom_arr, dtype="float64")
    return np.where(numer == 0, np.nan, (numer - denom) / numer * 100)

def safe_sum(series):
    return 0 if (series is None or len(series) == 0) else series.sum()

def normalize_id(x):
    if pd.isna(x):
        return ""
    return str(x).replace(",", "").strip()

def keep_existing(cols, available):
    return [c for c in cols if c in available]

def add_pct_col(df, out_col, numer_col, denom_col):
    if numer_col in df.columns and denom_col in df.columns:
        df[out_col] = discrepancy_pct(df[numer_col].values, df[denom_col].values)
        return True
    return False

def format_sheet(ws):
    """Auto-fit column widths and freeze the header row."""
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 60)
    ws.freeze_panes = "A2"

# ------------------ main ------------------
def make_comparison(input_xlsx: str):
    p = Path(input_xlsx)
    assert p.exists(), f"File not found: {p}"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_p = p.with_name(f"discrepancyReport_{timestamp}.xlsx")

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

    blis_cols = set(blis.columns) if blis is not None else set()
    dcm_cols = set(dcm.columns) if dcm is not None else set()
    celtra_cols = set(celtra.columns) if celtra is not None else set()

    dcm_has_imp_col = D_IMP in dcm_cols
    dcm_has_click_col = D_CLICK in dcm_cols
    dcm_has_invalid_click_col = D_INVALID in dcm_cols
    dcm_has_invalid_imp_col = D_INVALID_IMP in dcm_cols

    blis_has_req_col = B_REQ in blis_cols
    blis_has_shown_col = B_SHOWN in blis_cols
    blis_has_clicks_col = B_CLICKS in blis_cols
    blis_has_raw_col = B_RAW in blis_cols
    blis_has_win_col = B_WIN in blis_cols

    celtra_has_req_col = C_REQ in celtra_cols
    celtra_has_loaded_col = C_LOADED in celtra_cols
    celtra_has_rendered_col = C_RENDERED in celtra_cols
    celtra_has_clicks_col = C_CLICKS in celtra_cols

    # ensure mapping columns present
    for col in [M_BLIS, M_CELTRA, M_DCM]:
        if col not in mapping.columns:
            raise ValueError(f"Mapping must contain column: '{col}' (even if some values are empty).")

    # normalize mapping ids
    mapping[M_BLIS] = mapping[M_BLIS].astype(str).apply(normalize_id)
    mapping[M_CELTRA] = mapping[M_CELTRA].astype(str).apply(normalize_id)
    mapping[M_DCM] = mapping[M_DCM].astype(str).apply(normalize_id)

    # global flag: Celtra data is only relevant if the sheet exists AND mapping has Celtra IDs
    celtra_globally_present = celtra is not None and any(
        v for v in mapping[M_CELTRA].tolist() if v and v.lower() not in ("nan", "")
    )

    # ------------------ prepare Blis aggregation (creative + date) ------------------
    if blis is not None:
        if B_CREATIVE not in blis.columns:
            raise ValueError(f"Blis sheet must have '{B_CREATIVE}' column.")
        blis["__blis_key"] = blis[B_CREATIVE].astype(str).apply(normalize_id)
        blis["date"] = to_ymd(blis[B_DATE]) if B_DATE in blis.columns else to_ymd(blis.iloc[:, 0])
        for c in [B_REQ, B_SHOWN, B_CLICKS, B_RAW, B_WIN]:
            if c not in blis.columns:
                blis[c] = 0
            blis[c] = pd.to_numeric(blis[c], errors="coerce").fillna(0)

        blis_agg = (
            blis.groupby(["__blis_key", "date"], as_index=False)
            .agg({
                B_REQ: "sum",
                B_SHOWN: "sum",
                B_CLICKS: "sum",
                B_RAW: "sum",
                B_WIN: "sum",
            })
            .rename(columns={
                B_REQ: "Blis Requested Impressions",
                B_SHOWN: "Blis Shown Impressions",
                B_CLICKS: "Blis Clicks",
                B_RAW: "Blis Raw Clicks",
                B_WIN: "Blis Win Count",
            })
        )
        blis_agg["__blis_key"] = blis_agg["__blis_key"].astype(str).apply(normalize_id)
    else:
        blis_agg = pd.DataFrame(
            columns=[
                "__blis_key",
                "date",
                "Blis Requested Impressions",
                "Blis Shown Impressions",
                "Blis Clicks",
                "Blis Raw Clicks",
                "Blis Win Count",
            ]
        )

    # ------------------ prepare DCM aggregation (placement + date) ------------------
    if dcm is not None:
        if D_PLACEMENT not in dcm.columns:
            raise ValueError(f"DCM sheet must have '{D_PLACEMENT}' column.")
        dcm["__dcm_key"] = dcm[D_PLACEMENT].astype(str).apply(normalize_id)
        dcm["date"] = to_ymd(dcm[D_DATE]) if D_DATE in dcm.columns else to_ymd(dcm.iloc[:, 0])

        for c in [D_IMP, D_CLICK, D_INVALID, D_INVALID_IMP]:
            if c not in dcm.columns:
                dcm[c] = 0
            dcm[c] = pd.to_numeric(dcm[c], errors="coerce").fillna(0)

        dcm_agg = (
            dcm.groupby(["__dcm_key", "date"], as_index=False)
            .agg({
                D_IMP: "sum",
                D_CLICK: "sum",
                D_INVALID: "sum",
                D_INVALID_IMP: "sum",
            })
            .rename(columns={
                D_IMP: "DCM Impressions",
                D_CLICK: "DCM Clicks",
                D_INVALID: "DCM Invalid Clicks",
                D_INVALID_IMP: "DCM Invalid Impressions",
            })
        )
        dcm_agg["DCM Total Impressions"] = dcm_agg["DCM Impressions"] + dcm_agg["DCM Invalid Impressions"]
        dcm_agg["DCM Total Clicks"] = dcm_agg["DCM Clicks"] + dcm_agg["DCM Invalid Clicks"]
        dcm_agg["__dcm_key"] = dcm_agg["__dcm_key"].astype(str).apply(normalize_id)
    else:
        dcm_agg = pd.DataFrame(
            columns=[
                "__dcm_key",
                "date",
                "DCM Impressions",
                "DCM Clicks",
                "DCM Invalid Clicks",
                "DCM Invalid Impressions",
                "DCM Total Impressions",
                "DCM Total Clicks",
            ]
        )

    # ------------------ prepare Celtra aggregation (placement + date) ------------------
    if celtra is not None:
        if C_PLACEMENT not in celtra.columns:
            raise ValueError(f"Celtra sheet must have '{C_PLACEMENT}' column.")
        celtra["__celtra_key"] = celtra[C_PLACEMENT].astype(str).apply(normalize_id)
        celtra["date"] = to_ymd(celtra[C_DATE]) if C_DATE in celtra.columns else to_ymd(celtra.iloc[:, 0])

        for src_col in [C_REQ, C_LOADED, C_RENDERED]:
            if src_col not in celtra.columns:
                celtra[src_col] = 0
            celtra[src_col] = pd.to_numeric(celtra[src_col], errors="coerce").fillna(0)

        celtra["Celtra Clicks"] = pd.to_numeric(celtra[C_CLICKS], errors="coerce").fillna(0) if C_CLICKS in celtra.columns else 0

        celtra_agg = (
            celtra.groupby(["__celtra_key", "date"], as_index=False)
            .agg({
                C_REQ: "sum",
                C_LOADED: "sum",
                C_RENDERED: "sum",
                "Celtra Clicks": "sum",
            })
            .rename(columns={
                C_REQ: "Celtra Requested Impressions",
                C_LOADED: "Celtra Loaded Impressions",
                C_RENDERED: "Celtra Rendered Impressions",
            })
        )
        celtra_agg["__celtra_key"] = celtra_agg["__celtra_key"].astype(str).apply(normalize_id)
    else:
        celtra_agg = pd.DataFrame(
            columns=[
                "__celtra_key",
                "date",
                "Celtra Requested Impressions",
                "Celtra Loaded Impressions",
                "Celtra Rendered Impressions",
                "Celtra Clicks",
            ]
        )

    # ------------------ build mapping lookups ------------------
    dcm_to_celtras = mapping.groupby(M_DCM)[M_CELTRA].apply(
        lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower() != "nan"]))
    ).to_dict()
    dcm_to_blis = mapping.groupby(M_DCM)[M_BLIS].apply(
        lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower() != "nan"]))
    ).to_dict()
    celtra_to_blis = mapping.groupby(M_CELTRA)[M_BLIS].apply(
        lambda s: sorted(set([normalize_id(x) for x in s if x and x.lower() != "nan"]))
    ).to_dict()
    blis_to_celtra = mapping.drop_duplicates(subset=[M_BLIS]).set_index(M_BLIS)[M_CELTRA].to_dict()
    blis_to_dcm = mapping.drop_duplicates(subset=[M_BLIS]).set_index(M_BLIS)[M_DCM].to_dict()

    def choose_group_key(dcm_val, cel_val, blis_val):
        d = normalize_id(dcm_val)
        c = normalize_id(cel_val)
        b = normalize_id(blis_val)
        if d and d.lower() != "nan":
            return d
        if c and c.lower() != "nan":
            return c
        return b

    mapping["group_key"] = mapping.apply(lambda r: choose_group_key(r[M_DCM], r[M_CELTRA], r[M_BLIS]), axis=1)
    unique_group_keys = sorted(set(mapping["group_key"].tolist()))

    summary_rows = []
    with pd.ExcelWriter(out_p, engine="openpyxl") as writer:
        # mapping warnings
        dup_check = mapping.groupby(M_BLIS).agg({M_CELTRA: pd.Series.nunique, M_DCM: pd.Series.nunique})
        dup_warn = dup_check[(dup_check[M_CELTRA] > 1) | (dup_check[M_DCM] > 1)]
        if not dup_warn.empty:
            dup_warn.reset_index().to_excel(writer, sheet_name="Mapping_Warnings", index=False)
            format_sheet(writer.sheets["Mapping_Warnings"])

        for key in unique_group_keys:
            key_norm = normalize_id(key)

            is_dcm_group = key_norm in mapping[M_DCM].values
            is_celtra_group = key_norm in mapping[M_CELTRA].values and not is_dcm_group

            if is_dcm_group:
                dcm_id_for_group = key_norm
                celtra_ids = dcm_to_celtras.get(dcm_id_for_group, []) if celtra_globally_present else []
                blis_list = dcm_to_blis.get(dcm_id_for_group, [])
            elif is_celtra_group:
                cel_id_for_group = key_norm
                dcm_vals = mapping.loc[mapping[M_CELTRA] == cel_id_for_group, M_DCM].dropna().unique().tolist()
                dcm_id_for_group = normalize_id(dcm_vals[0]) if dcm_vals else ""
                celtra_ids = [cel_id_for_group] if celtra_globally_present else []
                blis_list = celtra_to_blis.get(cel_id_for_group, [])
            else:  # blis-group
                blis_creative = key_norm
                blis_list = [blis_creative]
                cel_id = normalize_id(blis_to_celtra.get(blis_creative, ""))
                dcm_id_for_group = normalize_id(blis_to_dcm.get(blis_creative, ""))
                celtra_ids = [cel_id] if (cel_id and celtra_globally_present) else []

            # Build dcm_frame
            if not dcm_agg.empty and dcm_id_for_group:
                dcm_frame = dcm_agg[dcm_agg["__dcm_key"] == dcm_id_for_group].copy()
            else:
                dcm_frame = pd.DataFrame(columns=dcm_agg.columns) if not dcm_agg.empty else pd.DataFrame()

            merged = pd.DataFrame()

            if is_dcm_group and celtra_ids:
                merged_rows = []

                def build_one_merged(cel_id, blis_list_for_cel):
                    if not celtra_agg.empty and cel_id:
                        cel_frame_local = celtra_agg[celtra_agg["__celtra_key"] == cel_id].copy()
                    else:
                        cel_frame_local = pd.DataFrame(columns=celtra_agg.columns) if not celtra_agg.empty else pd.DataFrame()

                    if not blis_agg.empty and blis_list_for_cel:
                        bf = blis_agg[blis_agg["__blis_key"].isin(blis_list_for_cel)].copy()
                        if not bf.empty:
                            blis_frame_local = bf.groupby("date", as_index=False).agg({
                                "Blis Requested Impressions": "sum",
                                "Blis Shown Impressions": "sum",
                                "Blis Clicks": "sum",
                                "Blis Raw Clicks": "sum",
                                "Blis Win Count": "sum",
                            })
                        else:
                            blis_frame_local = pd.DataFrame(columns=["date", "Blis Requested Impressions", "Blis Shown Impressions", "Blis Clicks", "Blis Raw Clicks", "Blis Win Count"])
                    else:
                        blis_frame_local = pd.DataFrame(columns=["date", "Blis Requested Impressions", "Blis Shown Impressions", "Blis Clicks", "Blis Raw Clicks", "Blis Win Count"])

                    date_sets_local = []
                    for df_local in (dcm_frame, cel_frame_local, blis_frame_local):
                        if isinstance(df_local, pd.DataFrame) and "date" in df_local.columns and not df_local["date"].dropna().empty:
                            date_sets_local.append(set(df_local["date"].dropna().astype(str).tolist()))
                        else:
                            date_sets_local.append(set())
                    all_dates_local = sorted(set().union(*date_sets_local))
                    if not all_dates_local:
                        return None

                    merged_local = pd.DataFrame({"date": all_dates_local})

                    if not dcm_frame.empty:
                        merged_local = merged_local.merge(
                            dcm_frame[[
                                "date",
                                "DCM Impressions",
                                "DCM Clicks",
                                "DCM Invalid Clicks",
                                "DCM Invalid Impressions",
                                "DCM Total Impressions",
                                "DCM Total Clicks",
                            ]],
                            on="date",
                            how="left",
                        )

                    if not blis_frame_local.empty:
                        merged_local = merged_local.merge(
                            blis_frame_local[[
                                "date",
                                "Blis Requested Impressions",
                                "Blis Shown Impressions",
                                "Blis Clicks",
                                "Blis Raw Clicks",
                                "Blis Win Count",
                            ]],
                            on="date",
                            how="left",
                        )

                    if not cel_frame_local.empty:
                        merged_local = merged_local.merge(
                            cel_frame_local[[
                                "date",
                                "Celtra Requested Impressions",
                                "Celtra Loaded Impressions",
                                "Celtra Rendered Impressions",
                                "Celtra Clicks",
                            ]],
                            on="date",
                            how="left",
                        )

                    merged_local["DCM Placement ID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                    merged_local["Celtra Placement ID"] = cel_id if cel_id else np.nan
                    merged_local["Blis Creative ID"] = ",".join(blis_list_for_cel) if blis_list_for_cel else np.nan

                    return merged_local

                for cel_id in celtra_ids:
                    blis_list_for_cel = celtra_to_blis.get(cel_id, [])
                    ml = build_one_merged(cel_id, blis_list_for_cel)
                    if ml is not None:
                        merged_rows.append(ml)

                if not merged_rows:
                    if not dcm_frame.empty and "date" in dcm_frame.columns and not dcm_frame["date"].dropna().empty:
                        merged_local = pd.DataFrame({"date": sorted(dcm_frame["date"].dropna().astype(str).tolist())})
                        merged_local = merged_local.merge(
                            dcm_frame[[
                                "date",
                                "DCM Impressions",
                                "DCM Clicks",
                                "DCM Invalid Clicks",
                                "DCM Invalid Impressions",
                                "DCM Total Impressions",
                                "DCM Total Clicks",
                            ]],
                            on="date",
                            how="left",
                        )
                        merged_local["Blis Requested Impressions"] = 0
                        merged_local["Blis Shown Impressions"] = 0
                        merged_local["Blis Clicks"] = 0
                        merged_local["Blis Raw Clicks"] = 0
                        merged_local["Blis Win Count"] = 0
                        merged_local["DCM Placement ID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                        merged_local["Celtra Placement ID"] = np.nan
                        merged_local["Blis Creative ID"] = np.nan
                        merged_rows.append(merged_local)

                merged = pd.concat(merged_rows, ignore_index=True, sort=False) if merged_rows else pd.DataFrame()

            else:
                if celtra_ids:
                    cel_frame = celtra_agg[celtra_agg["__celtra_key"].isin(celtra_ids)].copy()
                    if not cel_frame.empty:
                        cel_frame = cel_frame.groupby("date", as_index=False).agg({
                            "Celtra Requested Impressions": "sum",
                            "Celtra Loaded Impressions": "sum",
                            "Celtra Rendered Impressions": "sum",
                            "Celtra Clicks": "sum",
                        })
                else:
                    cel_frame = pd.DataFrame(columns=celtra_agg.columns) if not celtra_agg.empty else pd.DataFrame()

                if not blis_agg.empty and blis_list:
                    bf = blis_agg[blis_agg["__blis_key"].isin(blis_list)].copy()
                    if not bf.empty:
                        blis_frame = bf.groupby("date", as_index=False).agg({
                            "Blis Requested Impressions": "sum",
                            "Blis Shown Impressions": "sum",
                            "Blis Clicks": "sum",
                            "Blis Raw Clicks": "sum",
                            "Blis Win Count": "sum",
                        })
                    else:
                        blis_frame = pd.DataFrame(columns=["date", "Blis Requested Impressions", "Blis Shown Impressions", "Blis Clicks", "Blis Raw Clicks", "Blis Win Count"])
                else:
                    blis_frame = pd.DataFrame(columns=["date", "Blis Requested Impressions", "Blis Shown Impressions", "Blis Clicks", "Blis Raw Clicks", "Blis Win Count"])

                date_sets = []
                for df in (dcm_frame, cel_frame, blis_frame):
                    if isinstance(df, pd.DataFrame) and "date" in df.columns and not df["date"].dropna().empty:
                        date_sets.append(set(df["date"].dropna().astype(str).tolist()))
                    else:
                        date_sets.append(set())
                all_dates = sorted(set().union(*date_sets))

                if not all_dates:
                    merged = pd.DataFrame()
                else:
                    merged = pd.DataFrame({"date": all_dates})

                    if not dcm_frame.empty:
                        merged = merged.merge(
                            dcm_frame[[
                                "date",
                                "DCM Impressions",
                                "DCM Clicks",
                                "DCM Invalid Clicks",
                                "DCM Invalid Impressions",
                                "DCM Total Impressions",
                                "DCM Total Clicks",
                            ]],
                            on="date",
                            how="left",
                        )

                    if not blis_frame.empty:
                        merged = merged.merge(
                            blis_frame[[
                                "date",
                                "Blis Requested Impressions",
                                "Blis Shown Impressions",
                                "Blis Clicks",
                                "Blis Raw Clicks",
                                "Blis Win Count",
                            ]],
                            on="date",
                            how="left",
                        )

                    if not cel_frame.empty:
                        merged = merged.merge(
                            cel_frame[[
                                "date",
                                "Celtra Requested Impressions",
                                "Celtra Loaded Impressions",
                                "Celtra Rendered Impressions",
                                "Celtra Clicks",
                            ]],
                            on="date",
                            how="left",
                        )

                    merged["Blis Creative ID"] = ",".join(blis_list) if blis_list else np.nan
                    merged["DCM Placement ID"] = dcm_id_for_group if dcm_id_for_group else np.nan
                    merged["Celtra Placement ID"] = ",".join(celtra_ids) if celtra_ids else np.nan

            if merged.empty:
                summary_rows.append({
                    "group_key": key_norm,
                    "dcm_present": not dcm_frame.empty,
                    "celtra_present": len(celtra_ids) > 0,
                    "blis_present": len(blis_list) > 0,
                    "dcm_impressions_total": 0,
                    "blis_impressions_total": 0,
                    "celtra_rendered_total": 0,
                    "blis_vs_dcm_pct_total": np.nan,
                    "blis_vs_celtra_pct_total": np.nan,
                })
                continue

            out_numeric = [
                "DCM Impressions",
                "DCM Clicks",
                "DCM Invalid Clicks",
                "DCM Invalid Impressions",
                "DCM Total Impressions",
                "DCM Total Clicks",
                "Blis Requested Impressions",
                "Blis Shown Impressions",
                "Blis Clicks",
                "Blis Raw Clicks",
                "Blis Win Count",
                "Celtra Requested Impressions",
                "Celtra Loaded Impressions",
                "Celtra Rendered Impressions",
                "Celtra Clicks",
            ]
            for c in out_numeric:
                if c in merged.columns:
                    merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0)
                else:
                    merged[c] = 0

            merged["DCM Total Impressions"] = merged["DCM Impressions"] + merged["DCM Invalid Impressions"]
            merged["DCM Total Clicks"] = merged["DCM Clicks"] + merged["DCM Invalid Clicks"]

            has_blis_block = bool(blis_list)
            has_dcm_block = bool(dcm_id_for_group)
            has_celtra_block = bool(celtra_ids)

            base_cols = ["date"]
            if has_blis_block:
                base_cols.append("Blis Creative ID")
            if has_dcm_block:
                base_cols.append("DCM Placement ID")
            if has_celtra_block:
                base_cols.append("Celtra Placement ID")

            metric_cols = []
            if has_blis_block:
                metric_cols.extend(["Blis Requested Impressions", "Blis Shown Impressions", "Blis Clicks", "Blis Raw Clicks", "Blis Win Count"])
            if has_dcm_block:
                dcm_metric_cols = ["DCM Impressions", "DCM Clicks"]
                if dcm_has_invalid_click_col:
                    dcm_metric_cols.append("DCM Invalid Clicks")
                if dcm_has_invalid_imp_col:
                    dcm_metric_cols.append("DCM Invalid Impressions")
                dcm_metric_cols.extend(["DCM Total Impressions", "DCM Total Clicks"])
                metric_cols.extend(dcm_metric_cols)
            if has_celtra_block:
                celtra_metric_cols = ["Celtra Requested Impressions", "Celtra Loaded Impressions", "Celtra Rendered Impressions"]
                if celtra_has_clicks_col:
                    celtra_metric_cols.append("Celtra Clicks")
                metric_cols.extend(celtra_metric_cols)

            if has_blis_block and has_dcm_block:
                add_pct_col(merged, "Blis Requested Impressions vs DCM Impressions %", "Blis Requested Impressions", "DCM Impressions")
                add_pct_col(merged, "Blis Shown Impressions vs DCM Impressions %", "Blis Shown Impressions", "DCM Impressions")
                add_pct_col(merged, "Blis Requested Impressions vs DCM Total Impressions %", "Blis Requested Impressions", "DCM Total Impressions")
                add_pct_col(merged, "Blis Shown Impressions vs DCM Total Impressions %", "Blis Shown Impressions", "DCM Total Impressions")
                add_pct_col(merged, "Blis Clicks vs DCM Clicks %", "Blis Clicks", "DCM Clicks")
                add_pct_col(merged, "Blis Raw Clicks vs DCM Total Clicks %", "Blis Raw Clicks", "DCM Total Clicks")

            if has_blis_block and has_celtra_block:
                add_pct_col(merged, "Blis Requested Impressions vs Celtra Loaded Impressions %", "Blis Requested Impressions", "Celtra Loaded Impressions")
                add_pct_col(merged, "Blis Clicks vs Celtra Clicks %", "Blis Clicks", "Celtra Clicks")

            if has_celtra_block and has_dcm_block:
                add_pct_col(merged, "Celtra Loaded Impressions vs DCM Impressions %", "Celtra Loaded Impressions", "DCM Impressions")
                add_pct_col(merged, "Celtra Clicks vs DCM Clicks %", "Celtra Clicks", "DCM Clicks")

            pct_cols = [
                "Blis Requested Impressions vs DCM Impressions %",
                "Blis Shown Impressions vs DCM Impressions %",
                "Blis Requested Impressions vs DCM Total Impressions %",
                "Blis Shown Impressions vs DCM Total Impressions %",
                "Blis Clicks vs DCM Clicks %",
                "Blis Raw Clicks vs DCM Total Clicks %",
                "Blis Requested Impressions vs Celtra Loaded Impressions %",
                "Blis Clicks vs Celtra Clicks %",
                "Celtra Loaded Impressions vs DCM Impressions %",
                "Celtra Clicks vs DCM Clicks %",
            ]
            pct_cols = [c for c in pct_cols if c in merged.columns]

            final_cols = [c for c in base_cols + metric_cols + pct_cols if c in merged.columns]

            # add total row values
            numeric_for_sum = [
                c for c in merged.columns
                if c not in ("date", "Blis Creative ID", "DCM Placement ID", "Celtra Placement ID")
                and merged[c].dtype.kind in "biufc"
            ]
            totals = {c: safe_sum(merged[c]) for c in numeric_for_sum}

            if is_dcm_group and celtra_ids and not dcm_frame.empty:
                totals["DCM Impressions"] = dcm_frame["DCM Impressions"].sum() if "DCM Impressions" in dcm_frame.columns else 0
                totals["DCM Clicks"] = dcm_frame["DCM Clicks"].sum() if "DCM Clicks" in dcm_frame.columns else 0
                totals["DCM Invalid Clicks"] = dcm_frame["DCM Invalid Clicks"].sum() if "DCM Invalid Clicks" in dcm_frame.columns else 0
                totals["DCM Invalid Impressions"] = dcm_frame["DCM Invalid Impressions"].sum() if "DCM Invalid Impressions" in dcm_frame.columns else 0

            totals["DCM Total Impressions"] = totals.get("DCM Impressions", 0) + totals.get("DCM Invalid Impressions", 0)
            totals["DCM Total Clicks"] = totals.get("DCM Clicks", 0) + totals.get("DCM Invalid Clicks", 0)

            def pct_total(numer_total, denom_total):
                return round((numer_total - denom_total) / numer_total * 100, 2) if numer_total != 0 else np.nan

            if "Blis Requested Impressions" in merged.columns and "DCM Impressions" in merged.columns:
                totals["Blis Requested Impressions vs DCM Impressions %"] = pct_total(totals.get("Blis Requested Impressions", 0), totals.get("DCM Impressions", 0))
            if "Blis Shown Impressions" in merged.columns and "DCM Impressions" in merged.columns:
                totals["Blis Shown Impressions vs DCM Impressions %"] = pct_total(
                    totals.get("Blis Shown Impressions", 0),
                    totals.get("DCM Impressions", 0)
                )
            if "Blis Requested Impressions" in merged.columns and "DCM Total Impressions" in merged.columns:
                totals["Blis Requested Impressions vs DCM Total Impressions %"] = pct_total(totals.get("Blis Requested Impressions", 0), totals.get("DCM Total Impressions", 0))
            if "Blis Shown Impressions" in merged.columns and "DCM Total Impressions" in merged.columns:
                totals["Blis Shown Impressions vs DCM Total Impressions %"] = pct_total(totals.get("Blis Shown Impressions", 0), totals.get("DCM Total Impressions", 0))
            if "Blis Clicks" in merged.columns and "DCM Clicks" in merged.columns:
                totals["Blis Clicks vs DCM Clicks %"] = pct_total(totals.get("Blis Clicks", 0), totals.get("DCM Clicks", 0))
            if "Blis Raw Clicks" in merged.columns and "DCM Total Clicks" in merged.columns:
                totals["Blis Raw Clicks vs DCM Total Clicks %"] = pct_total(totals.get("Blis Raw Clicks", 0), totals.get("DCM Total Clicks", 0))
            if "Blis Requested Impressions" in merged.columns and "Celtra Loaded Impressions" in merged.columns:
                totals["Blis Requested Impressions vs Celtra Loaded Impressions %"] = pct_total(totals.get("Blis Requested Impressions", 0), totals.get("Celtra Loaded Impressions", 0))
            if "Blis Clicks" in merged.columns and "Celtra Clicks" in merged.columns:
                totals["Blis Clicks vs Celtra Clicks %"] = pct_total(totals.get("Blis Clicks", 0), totals.get("Celtra Clicks", 0))
            if "Celtra Loaded Impressions" in merged.columns and "DCM Impressions" in merged.columns:
                totals["Celtra Loaded Impressions vs DCM Impressions %"] = pct_total(totals.get("Celtra Loaded Impressions", 0), totals.get("DCM Impressions", 0))
            if "Celtra Clicks" in merged.columns and "DCM Clicks" in merged.columns:
                totals["Celtra Clicks vs DCM Clicks %"] = pct_total(totals.get("Celtra Clicks", 0), totals.get("DCM Clicks", 0))

            totals["Blis Creative ID"] = ",".join(blis_list) if blis_list else np.nan
            totals["DCM Placement ID"] = dcm_id_for_group if dcm_id_for_group else np.nan
            totals["Celtra Placement ID"] = ",".join(celtra_ids) if (not is_dcm_group and celtra_ids) else (np.nan if is_dcm_group else np.nan)
            totals["date"] = "Grand Total"

            total_row = pd.DataFrame([totals])
            total_row = total_row.reindex(columns=merged.columns, fill_value=0)

            for c in merged.columns:
                if "%" in c and c in total_row.columns:
                    v = total_row.at[0, c]
                    total_row.at[0, c] = round(v, 2) if pd.notna(v) else v

            # drop data rows where all metrics are zero (keep Grand Total)
            id_cols = {"date", "Blis Creative ID", "DCM Placement ID", "Celtra Placement ID"}
            metric_only_cols = [c for c in final_cols if c not in id_cols]
            data_rows = merged[final_cols].copy()
            if metric_only_cols:
                non_zero_mask = data_rows[metric_only_cols].fillna(0).abs().sum(axis=1) > 0
                data_rows = data_rows[non_zero_mask]

            merged_out = pd.concat([data_rows, total_row[final_cols]], ignore_index=True, sort=False)

            sheet_name = f"Group_{key_norm}"[:31]
            merged_out.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer.sheets[sheet_name])

            summary_rows.append({
                "group_key": key_norm,

                "dcm_present": not dcm_frame.empty,
                "celtra_present": len(celtra_ids) > 0,
                "blis_present": len(blis_list) > 0,

                "blis_requested_total": totals.get("Blis Requested Impressions", 0),
                "blis_shown_total": totals.get("Blis Shown Impressions", 0),
                "blis_click_total": totals.get("Blis Clicks", 0),
                "blis_raw_click_total": totals.get("Blis Raw Clicks", 0),

                "dcm_impression_total": totals.get("DCM Impressions", 0),
                "dcm_invalid_impression_total": totals.get("DCM Invalid Impressions", 0),
                "dcm_total_impression": totals.get("DCM Total Impressions", 0),

                "dcm_click_total": totals.get("DCM Clicks", 0),
                "dcm_invalid_click_total": totals.get("DCM Invalid Clicks", 0),
                "dcm_total_click": totals.get("DCM Total Clicks", 0),

                "celtra_loaded_total": totals.get("Celtra Loaded Impressions", 0),
                "celtra_rendered_total": totals.get("Celtra Rendered Impressions", 0),
                "celtra_click_total": totals.get("Celtra Clicks", 0),

                "blis_req_vs_dcm_imp_%": totals.get("Blis Requested Impressions vs DCM Impressions %"),
                "blis_shown_vs_dcm_imp_%": totals.get("Blis Shown Impressions vs DCM Impressions %"),

                "blis_req_vs_dcm_total_%": totals.get("Blis Requested Impressions vs DCM Total Impressions %"),
                "blis_shown_vs_dcm_total_%": totals.get("Blis Shown Impressions vs DCM Total Impressions %"),

                "blis_click_vs_dcm_click_%": totals.get("Blis Clicks vs DCM Clicks %"),
                "blis_raw_vs_dcm_total_click_%": totals.get("Blis Raw Clicks vs DCM Total Clicks %"),

                "blis_vs_celtra_imp_%": totals.get("Blis Requested Impressions vs Celtra Loaded Impressions %"),
                "blis_vs_celtra_click_%": totals.get("Blis Clicks vs Celtra Clicks %"),
            })

        summary_df = pd.DataFrame(summary_rows)

        # drop Celtra columns from Summary when no group has Celtra data
        if not celtra_globally_present or not any(r.get("celtra_present", False) for r in summary_rows):
            celtra_summary_cols = [
                "celtra_present",
                "celtra_loaded_total",
                "celtra_rendered_total",
                "celtra_click_total",
                "blis_vs_celtra_imp_%",
                "blis_vs_celtra_click_%",
            ]
            summary_df = summary_df.drop(columns=[c for c in celtra_summary_cols if c in summary_df.columns])

        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        format_sheet(writer.sheets["Summary"])

    print("Done. Output written to:", out_p)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python descrepancy_automation_script.py descrepancyTemplate.xlsx")
        sys.exit(1)
    make_comparison(sys.argv[1])
