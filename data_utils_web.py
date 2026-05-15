"""
data_utils_web.py — Data-processing helpers for the Streamlit web app.
Identical to data_utils.py but without the Qt matplotlib backend setting,
and load_csv accepts a file-like object OR a path string.
"""

import io
import json
import re
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.signal import savgol_filter
import matplotlib
matplotlib.use('Agg')   # non-interactive backend for web


# ── CSV loading ───────────────────────────────────────────────────────────────
def load_csv(source) -> pd.DataFrame:
    """
    Load a displacement/load CSV from a path string or file-like object.
    Returns a DataFrame with columns 'displacement' and 'load'.
    """
    if isinstance(source, (str, Path)):
        df = pd.read_csv(source, skipinitialspace=True)
    else:
        df = pd.read_csv(io.BytesIO(source.read()) if hasattr(source, 'read') else source,
                         skipinitialspace=True)
    df.columns = [c.strip() for c in df.columns]
    disp_col = load_col = None

    # Pass 1: prefer 'digital position' or 'displacement…digital'
    for col in df.columns:
        cl = col.lower()
        if 'digital position' in cl or ('displacement' in cl and 'digital' in cl):
            if disp_col is None: disp_col = col
        elif cl.startswith('load'):
            if load_col is None: load_col = col

    # Pass 2: any column containing 'displacement'
    if disp_col is None:
        for col in df.columns:
            if 'displacement' in col.lower() and 'total' not in col.lower():
                disp_col = col; break

    # Pass 3: any column containing 'position'
    if disp_col is None:
        for col in df.columns:
            if 'position' in col.lower() and 'total' not in col.lower():
                disp_col = col; break

    # Pass 4: first predominantly-numeric column
    if disp_col is None:
        for col in df.columns:
            if pd.to_numeric(df[col], errors='coerce').notna().sum() > len(df) // 2:
                disp_col = col; break

    # Pass 5: load fallback — 'load', 'force', 'charge'
    if load_col is None:
        for col in df.columns:
            cl = col.lower()
            if 'load' in cl or 'force' in cl or 'charge' in cl:
                load_col = col; break

    if not disp_col or not load_col:
        raise ValueError(
            f"Cannot find displacement/load columns.\nFound: {list(df.columns)}")

    return pd.DataFrame({
        'displacement': pd.to_numeric(df[disp_col], errors='coerce'),
        'load':         pd.to_numeric(df[load_col],  errors='coerce'),
    }).dropna()


# ── Signal processing ─────────────────────────────────────────────────────────
def smooth_data(y: np.ndarray, window: int) -> np.ndarray:
    w = window if window % 2 == 1 else window + 1
    w = max(w, 5)
    return savgol_filter(y, w, 3) if len(y) >= w else y.copy()


def calc_stats(y: np.ndarray) -> dict:
    if len(y) == 0:
        return {'mean': 0.0, 'max': 0.0, 'min': 0.0, 'std': 0.0}
    return {
        'mean': float(np.mean(y)),
        'max':  float(np.max(y)),
        'min':  float(np.min(y)),
        'std':  float(np.std(y)),
    }


# ── XLSX export ───────────────────────────────────────────────────────────────
def export_tests_to_xlsx_bytes(tests, db) -> bytes:
    """Return Excel file as bytes (for st.download_button)."""
    sf_map = {sf['id']: sf for sf in db.get_source_folders()}
    rows = []
    for t in tests:
        a = db.get_analysis(t['id'])
        if not a:
            continue
        sf        = sf_map.get(t.get('source_folder_id'), {})
        exp_orig  = sf.get('name', '')
        exp_disp  = sf.get('display_name') or exp_orig
        test_orig = t.get('test_folder_name', '')
        test_disp = t.get('display_name') or test_orig
        row = {
            'Experiment':            exp_disp,
            'Experiment (original)': exp_orig,
            'Test Name':             test_disp,
            'Test Name (original)':  test_orig,
            'Test Number':           t.get('test_number', ''),
            'Whole Mean Raw':        a.get('whole_mean'),
            'Whole Max Raw':         a.get('whole_max'),
            'Whole Min Raw':         a.get('whole_min'),
            'Whole Std Raw':         a.get('whole_std'),
            'Whole Mean Smooth':     a.get('whole_mean_smooth'),
            'Whole Max Smooth':      a.get('whole_max_smooth'),
            'Whole Min Smooth':      a.get('whole_min_smooth'),
            'Whole Std Smooth':      a.get('whole_std_smooth'),
        }
        regions = json.loads(a.get('regions_json', '[]') or '[]')
        for i, reg in enumerate(regions):
            pfx = f"R{i+1}_{'C' if reg.get('consistent') else 'X'}"
            row[f"{pfx}_start_mm"] = reg.get('xmin', '')
            row[f"{pfx}_end_mm"]   = reg.get('xmax', '')
            try:
                raw = db.get_csv_data(t['id'])
                df  = load_csv(io.BytesIO(raw))
                x, y = df['displacement'].values, df['load'].values
                w    = int(a.get('smoothing_window', 51) or 51)
                ys   = smooth_data(y, w)
                mask = (x >= reg['xmin']) & (x <= reg['xmax'])
                sr   = calc_stats(y[mask])
                ss   = calc_stats(ys[mask])
                for k in ['mean', 'max', 'min', 'std']:
                    row[f"{pfx}_{k}_raw"]    = sr[k]
                    row[f"{pfx}_{k}_smooth"] = ss[k]
            except Exception:
                pass
        rows.append(row)

    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    return buf.getvalue()
