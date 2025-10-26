# automation_script_fixed.py
"""
Updated script: aggregates per-date properly and fills zeros when Blis is missing dates.
Usage:
    python automation_script_fixed.py /path/to/sample_discrepancy_fixed.xlsx
"""
import sys
from pathlib import Path
import pandas as pd
import numpy as np

def make_comparison(wb_path: str):
    wb_path = Path(wb_path)
    assert wb_path.exists(), f"File not found: {wb_path}"
    out_path = wb_path.with_name(wb_path.stem + "_comparison_fixed.xlsx")

    # Read sheets
    mapping = pd.read_excel(wb_path, sheet_name="Mapping")
    blis = pd.read_excel(wb_path, sheet_name="Blis")
    dcm = pd.read_excel(wb_path, sheet_name="DCM")

    # --- Normalize date columns to YYYY-MM-DD strings (stable for merging) ---
    blis['date'] = pd.to_datetime(blis['date'], errors='coerce').dt.strftime('%Y-%m-%d')
    dcm['date'] = pd.to_datetime(dcm['date'], errors='coerce').dt.strftime('%Y-%m-%d')

    # --- Ensure numeric columns exist and are numeric, fill NaN with 0 ---
    blis_numeric_cols = ['impressions', 'total_win_count', 'shownImpression', 'clicks', 'raw_clicks']
    for c in blis_numeric_cols:
        if c in blis.columns:
            blis[c] = pd.to_numeric(blis[c], errors='coerce').fillna(0)
        else:
            # If column missing in Blis, create it with zeros so aggregation code stays consistent
            blis[c] = 0

    # In DCM, ensure these columns exist (invalid_clicks is optional)
    dcm_numeric_cols = ['impressions', 'clicks', 'invalid_clicks']
    for c in dcm_numeric_cols:
        if c in dcm.columns:
            dcm[c] = pd.to_numeric(dcm[c], errors='coerce').fillna(0)
        else:
            dcm[c] = 0

    # --- Validate mapping: check if any creative maps to multiple placements ---
    dup = mapping.groupby('blisCreativeID')['placementID'].nunique()
    problematic = dup[dup > 1]

    # Pre-aggregate Blis at (date, creative) level just in case of duplicate rows
    blis_agg = blis.groupby(['date', 'blisCreativeID'], as_index=False).agg({
        'impressions': 'sum',
        'total_win_count': 'sum',
        'shownImpression': 'sum',
        'clicks': 'sum',
        'raw_clicks': 'sum'
    })

    # Build creative -> placement map (if duplicate creative entries in mapping, keep first)
    creative_to_placement = mapping.drop_duplicates(subset=['blisCreativeID']).set_index('blisCreativeID')['placementID'].to_dict()

    # Map creatives to placement and drop creatives not present in mapping (they can't be compared)
    blis_agg['placementID'] = blis_agg['blisCreativeID'].map(creative_to_placement)
    blis_mapped = blis_agg[~blis_agg['placementID'].isna()].copy()
    # Ensure placementID is same type as in mapping/DCM
    blis_mapped['placementID'] = blis_mapped['placementID'].astype(mapping['placementID'].dtype)

    # Aggregate Blis to placement + date (sum across creatives)
    blis_by_placement = blis_mapped.groupby(['placementID', 'date'], as_index=False).agg({
        'impressions': 'sum',
        'total_win_count': 'sum',
        'shownImpression': 'sum',
        'clicks': 'sum',
        'raw_clicks': 'sum'
    }).rename(columns={
        'impressions': 'blis_impressions',
        'total_win_count': 'blis_total_win_count',
        'shownImpression': 'blis_shownImpression',
        'clicks': 'blis_clicks',
        'raw_clicks': 'blis_raw_clicks'
    })

    # Rename DCM numeric columns for clarity
    dcm = dcm.rename(columns={'impressions': 'dcm_impressions', 'clicks': 'dcm_clicks', 'invalid_clicks': 'dcm_invalid_clicks'})

    placements = sorted(mapping['placementID'].unique().tolist())
    summary_rows = []

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # If problematic mapping exists, write warnings sheet
        if not problematic.empty:
            warn_df = mapping[mapping['blisCreativeID'].isin(problematic.index)].sort_values('blisCreativeID')
            warn_df.to_excel(writer, sheet_name='Mapping_Warnings', index=False)
            print("Warning: Some creatives map to multiple placements. See 'Mapping_Warnings' sheet in output.")

        for pid in placements:
            pid_val = pid  # keep original type
            # DCM rows for this placement
            # --- inside your for pid in placements: loop, replace the current block with this ---

# Filter original DCM rows for this placement and AGGREGATE by date (important!)
            dcm_pid_raw = dcm[dcm['placementID'] == pid_val][['date', 'dcm_impressions', 'dcm_clicks', 'dcm_invalid_clicks']].copy()
            dcm_pid = dcm_pid_raw.groupby('date', as_index=False).agg({
                'dcm_impressions': 'sum',
                'dcm_clicks': 'sum',
                'dcm_invalid_clicks': 'sum'
            })

            # Ensure blis_pid is aggregated by date as well (safe-guard)
            blis_pid_raw = blis_by_placement[blis_by_placement['placementID'] == pid_val][
                ['date', 'blis_impressions', 'blis_shownImpression', 'blis_clicks', 'blis_raw_clicks', 'blis_total_win_count']
            ].copy()
            blis_pid = blis_pid_raw.groupby('date', as_index=False).agg({
                'blis_impressions': 'sum',
                'blis_shownImpression': 'sum',
                'blis_clicks': 'sum',
                'blis_raw_clicks': 'sum',
                'blis_total_win_count': 'sum'
            })

            # Build full date list (union of DCM dates and Blis dates for this placement)
            all_dates = sorted(set(dcm_pid['date'].tolist()) | set(blis_pid['date'].tolist()))

            if not all_dates:
                continue

            dates_df = pd.DataFrame({'date': all_dates})

            # Merge aggregated DCM and aggregated BLIS into dates_df
            merged = dates_df.merge(dcm_pid, on='date', how='left').merge(blis_pid, on='date', how='left')

            # Fill NaNs for numeric columns with 0
            numeric_cols = ['dcm_impressions', 'dcm_clicks', 'dcm_invalid_clicks',
                            'blis_impressions', 'blis_shownImpression', 'blis_clicks', 'blis_raw_clicks', 'blis_total_win_count']
            for c in numeric_cols:
                if c in merged.columns:
                    merged[c] = pd.to_numeric(merged[c], errors='coerce').fillna(0)

            # Add placement and diff / pct columns as before
            merged['placementID'] = pid_val
            merged['imp_diff'] = merged['dcm_impressions'] - merged['blis_impressions']
            merged['click_diff'] = merged['dcm_clicks'] - merged['blis_clicks']

            def pct_safe(blis_col, dcm_col):
                return np.where(merged[dcm_col] == 0, np.nan, (merged[blis_col] - merged[dcm_col]) / merged[dcm_col] * 100)

            merged['blis_imp_vs_DCM_imp_pct'] = pct_safe('blis_impressions', 'dcm_impressions')
            merged['blis_shownImp_vs_DCM_imp_pct'] = pct_safe('blis_shownImpression', 'dcm_impressions')
            merged['blis_click_vs_DCM_click_pct'] = pct_safe('blis_clicks', 'dcm_clicks')
            merged['blis_raw_click_vs_DCM_invalid_click_pct'] = pct_safe('blis_raw_clicks', 'dcm_invalid_clicks')

            # reorder / pick columns and write sheet as you already do


            # Reorder columns for readability (optional)
            cols_order = ['date', 'placementID',
                          'dcm_impressions', 'blis_impressions', 'imp_diff', 'blis_imp_vs_DCM_imp_pct',
                          'dcm_clicks', 'blis_clicks', 'click_diff', 'blis_click_vs_DCM_click_pct',
                          'dcm_invalid_clicks', 'blis_raw_clicks', 'blis_raw_click_vs_DCM_invalid_click_pct',
                          'blis_shownImpression', 'blis_total_win_count'
                          ]
            # Keep only those that exist
            cols_order = [c for c in cols_order if c in merged.columns]
            merged = merged[cols_order]

            # Write per-placement sheet
            

            # --- Insert this after you've computed `merged` and done `merged = merged[cols_order]`
            # and BEFORE: merged.to_excel(writer, sheet_name=sheet_name, index=False)

            # 1) find numeric columns in merged (fast)
            numeric_cols = [c for c in merged.columns if merged[c].dtype.kind in 'biufc']

            # 2) compute totals once (vectorized)
            totals = merged[numeric_cols].sum().to_dict()

            # 3) add diffs totals (dcm - blis)
            totals['imp_diff'] = totals.get('dcm_impressions', 0) - totals.get('blis_impressions', 0)
            totals['click_diff'] = totals.get('dcm_clicks', 0) - totals.get('blis_clicks', 0)

            # 4) helper for safe percent totals (rounded)
            def pct_total(blis_total, dcm_total):
                return round((blis_total - dcm_total) / dcm_total * 100, 2) if dcm_total != 0 else np.nan

            totals['blis_imp_vs_DCM_imp_pct'] = pct_total(totals.get('blis_impressions', 0), totals.get('dcm_impressions', 0))
            totals['blis_shownImp_vs_DCM_imp_pct'] = pct_total(totals.get('blis_shownImpression', 0), totals.get('dcm_impressions', 0))
            totals['blis_click_vs_DCM_click_pct'] = pct_total(totals.get('blis_clicks', 0), totals.get('dcm_clicks', 0))
            totals['blis_raw_click_vs_DCM_invalid_click_pct'] = pct_total(totals.get('blis_raw_clicks', 0), totals.get('dcm_invalid_clicks', 0))

            # 5) add placement and label for the totals row
            totals['placementID'] = pid_val
            totals['date'] = 'Grand Total'

            # 6) build a one-row DataFrame and align its columns to merged
            total_row_df = pd.DataFrame([totals])
            total_row_df = total_row_df.reindex(columns=merged.columns, fill_value=0)

            # 7) optionally keep percent columns as floats / round them
            pct_cols = [c for c in merged.columns if 'pct' in c]
            for c in pct_cols:
                if c in total_row_df.columns:
                    # keep NaN if created, otherwise round
                    total_row_df[c] = total_row_df[c].apply(lambda v: round(v, 2) if pd.notna(v) else v)

            # 8) append total row (fast)
            merged = pd.concat([merged, total_row_df], ignore_index=True)

            sheet_name = f"{pid_val}"[:31]
            # 9) Now write the sheet
            merged.to_excel(writer, sheet_name=sheet_name, index=False)

            # 10) Use the precomputed totals (not re-summing merged) to build summary row
            summary_rows.append({
                'placementID': pid_val,
                'dcm_impressions_total': totals.get('dcm_impressions', 0),
                'dcm_clicks_total': totals.get('dcm_clicks', 0),
                'dcm_invalid_clicks_total': totals.get('dcm_invalid_clicks', 0),
                'blis_impressions_total': totals.get('blis_impressions', 0),
                'blis_total_win_count_total': totals.get('blis_total_win_count', 0),
                'blis_shownImpression_total': totals.get('blis_shownImpression', 0),
                'blis_clicks_total': totals.get('blis_clicks', 0),
                'blis_raw_clicks_total': totals.get('blis_raw_clicks', 0),
                'blis_imp_vs_DCM_imp_pct_total': totals.get('blis_imp_vs_DCM_imp_pct'),
                'blis_click_vs_DCM_click_pct_total': totals.get('blis_click_vs_DCM_click_pct'),
            })

            

        # Write summary sheet
        summary = pd.DataFrame(summary_rows)
        if not summary.empty:
            summary.to_excel(writer, sheet_name='Summary', index=False)

    print(f"Comparison workbook saved to: {out_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python automation_script_fixed.py /path/to/sample_discrepancy_fixed.xlsx")
    else:
        make_comparison(sys.argv[1])
