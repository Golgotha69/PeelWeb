"""
streamlit_app.py — PeelAnalyzer Web (Streamlit edition) v1.6

Pages:
  🏠  Home / Import          — individual CSV, multi-CSV, or ZIP upload
  🔬  Analyse Tests          — single-test view + batch (shared regions)
  🔁  Sequential Analysis    — one-by-one: each test shown in turn, user sets
                               regions, saves & advances; Excel at the end
  📊  Figure Designer        — conditions, overlays, publication figures
  ⚙   Settings               — colours, units, smoothing default
"""

import io
import json
import re
import zipfile
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import streamlit as st

from data_utils_web import load_csv, smooth_data, calc_stats, export_tests_to_xlsx_bytes
from database_web import WebDatabase

# ─────────────────────────────────────────────────────────────────────────────
#  Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PeelAnalyzer Web",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
#  Session-state bootstrap
# ─────────────────────────────────────────────────────────────────────────────
def _init_state():
    defaults = {
        'db':                WebDatabase(),
        'load_unit':         'N',
        'default_smoothing': 51,
        'graph_bg':          '#13151a',
        'graph_raw_color':   '#3b82f6',
        'graph_smooth_color':'#60a5fa',
        'graph_grid_color':  '#374151',
        # sequential-analysis state machine
        'seq_active':        False,
        'seq_tests':         [],
        'seq_idx':           0,
        'seq_regions':       [],   # regions for the CURRENT test
        'seq_smoothing':     51,
        'seq_done':          False,
        'seq_folder_id':     None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()
db: WebDatabase = st.session_state.db

# ─────────────────────────────────────────────────────────────────────────────
#  Shared helpers
# ─────────────────────────────────────────────────────────────────────────────
REGION_COLORS = ['#f59e0b','#10b981','#ec4899','#8b5cf6','#06b6d4','#ef4444']

def _style_ax(ax, fig, title='', xlabel=None, ylabel=None):
    bg = st.session_state.graph_bg
    ax.set_facecolor(bg); fig.patch.set_facecolor(bg)
    ax.tick_params(colors='#9ca3af', labelsize=9)
    for sp in ['bottom','left']:  ax.spines[sp].set_color('#4b5563')
    for sp in ['top','right']:    ax.spines[sp].set_visible(False)
    ax.set_xlabel(xlabel or 'Displacement (mm)', color='#9ca3af', fontsize=10)
    ax.set_ylabel(ylabel or f"Load ({st.session_state.load_unit})",
                  color='#9ca3af', fontsize=10)
    ax.grid(True, color=st.session_state.graph_grid_color, linewidth=0.5, alpha=0.6)
    if title:
        ax.set_title(title, color='#f9fafb', fontsize=11, fontweight='bold')

def _load_test_data(t: dict):
    raw = db.get_csv_data(t['id'])
    if not raw: return None, None
    try:
        df = load_csv(io.BytesIO(raw))
        return df['displacement'].values, df['load'].values
    except Exception as e:
        st.warning(f"Cannot load {t.get('test_folder_name','?')}: {e}")
        return None, None

def _test_label(t: dict) -> str:
    return t.get('display_name') or t['test_folder_name']

def _folder_label(sf: dict) -> str:
    return sf.get('display_name') or sf['name']

def _draw_test_figure(x, y, ys, regions, title='', figsize=(9, 4)):
    fig, ax = plt.subplots(figsize=figsize)
    _style_ax(ax, fig, title=title)
    ax.plot(x, y,  color=st.session_state.graph_raw_color,
            lw=0.8, alpha=0.4, label='Raw')
    ax.plot(x, ys, color=st.session_state.graph_smooth_color,
            lw=1.8, label='Smoothed')
    for i, reg in enumerate(regions):
        c = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
        ax.axvspan(reg['xmin'], reg['xmax'], alpha=0.15, color=c, zorder=3)
        ax.axvline(reg['xmin'], color=c, lw=1, ls='--', alpha=0.8)
        ax.axvline(reg['xmax'], color=c, lw=1, ls='--', alpha=0.8)
    if regions:
        ax.legend(fontsize=8, facecolor='#1a1d23', labelcolor='#e8eaf0',
                  edgecolor='#374151', framealpha=0.9)
    else:
        ax.legend(fontsize=8, facecolor='#1a1d23', labelcolor='#e8eaf0',
                  edgecolor='#374151', framealpha=0.9)
    fig.tight_layout(pad=1.5)
    return fig

def _try_import_csv(raw_bytes: bytes, filename: str) -> bool:
    """Return True if load_csv accepts these bytes."""
    try:
        load_csv(io.BytesIO(raw_bytes))
        return True
    except Exception:
        return False

# ─────────────────────────────────────────────────────────────────────────────
#  ZIP import helper
# ─────────────────────────────────────────────────────────────────────────────
def _import_zip(zip_bytes: bytes, exp_name: str) -> tuple[int, list[str]]:
    """
    Parse a ZIP where each subfolder contains exactly one CSV (plus other
    files that are ignored).  Returns (count_imported, list_of_errors).

    Folder structure accepted:
      experiment/
        test_01/  ← becomes test label
          data.csv
          data.txt   ← ignored
        test_02/
          data.Stop.csv
    OR flat (no subfolders):
      experiment/
        test_01.csv
        test_02.csv
    """
    errors = []
    sfid   = db.add_source_folder(exp_name.strip() or "Unnamed")
    count  = 0

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = zf.namelist()

        # Build a map: folder_path → [csv_members]
        # Normalise paths: strip leading component if all share the same root
        # (e.g. when zipped as "experiment/test1/file.csv")
        parts_list = [n.replace('\\', '/') for n in names if not n.endswith('/')]

        # Detect common root prefix (zip made from a single top-level folder)
        roots = {p.split('/')[0] for p in parts_list if '/' in p}
        if len(roots) == 1 and all('/' in p for p in parts_list):
            root_prefix = roots.pop() + '/'
            parts_list  = [p[len(root_prefix):] for p in parts_list]
        else:
            root_prefix = ''

        # Group CSVs by their immediate parent folder
        from collections import defaultdict
        folder_csv: dict = defaultdict(list)
        flat_csvs = []

        for p in parts_list:
            if not p: continue
            segs = p.split('/')
            if len(segs) == 1:
                # Flat file at root level
                if p.lower().endswith('.csv'):
                    flat_csvs.append(p)
            else:
                # File inside a subfolder
                folder_name = segs[0]
                fname       = segs[-1]
                if fname.lower().endswith('.csv'):
                    folder_csv[folder_name].append(p)

        def _pick_best_csv(csv_paths):
            """Prefer *.Stop.csv or *.stop.csv; fallback to first."""
            stops = [p for p in csv_paths if 'stop' in p.lower()]
            return stops[0] if stops else csv_paths[0]

        # Import subfolder tests
        test_num = 1
        for folder_name in sorted(folder_csv.keys()):
            csv_paths = folder_csv[folder_name]
            best      = _pick_best_csv(csv_paths)
            zip_key   = root_prefix + best if root_prefix else best
            try:
                raw = zf.read(zip_key)
                if not _try_import_csv(raw, best):
                    errors.append(f"{folder_name}: CSV could not be parsed")
                    continue
                # Extract test number from folder name if possible
                m = re.search(r'(\d+)', folder_name)
                tnum = int(m.group(1)) if m else test_num
                db.add_test(sfid, folder_name, tnum, raw, best)
                count    += 1
                test_num += 1
            except Exception as e:
                errors.append(f"{folder_name}: {e}")

        # Import flat CSVs (no subfolders)
        for i, fname in enumerate(sorted(flat_csvs), 1):
            zip_key = root_prefix + fname if root_prefix else fname
            try:
                raw = zf.read(zip_key)
                if not _try_import_csv(raw, fname):
                    errors.append(f"{fname}: could not be parsed")
                    continue
                tname = re.sub(r'\.(csv|CSV)$', '', fname)
                m     = re.search(r'(\d+)', tname)
                tnum  = int(m.group(1)) if m else i
                db.add_test(sfid, tname, tnum, raw, fname)
                count += 1
            except Exception as e:
                errors.append(f"{fname}: {e}")

    return count, errors

# ─────────────────────────────────────────────────────────────────────────────
#  Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔬 PeelAnalyzer Web")
    st.markdown("---")

    page = st.radio("Navigation", [
        "🏠  Home / Import",
        "🔬  Analyse Tests",
        "🔁  Sequential Analysis",
        "📊  Figure Designer",
        "⚙   Settings",
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("### 💾 Session Database")
    db_up = st.file_uploader("Load previous session (.db)", type=['db'],
                              key='db_upload')
    if db_up is not None:
        if st.button("📥 Load this database"):
            st.session_state.db = WebDatabase(db_up.read())
            db = st.session_state.db
            st.success("Database loaded!")
            st.rerun()

    st.download_button(
        "📤 Download session (.db)",
        data=db.to_bytes(),
        file_name="peel_analyzer_session.db",
        mime="application/octet-stream",
    )
    st.markdown("---")
    folders = db.get_source_folders()
    st.caption(f"{len(folders)} folder(s)  •  {len(db.get_all_tests())} test(s)")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: HOME / IMPORT
# ─────────────────────────────────────────────────────────────────────────────
if "Home" in page:
    st.title("🏠  Home — Import")

    st.info(
        "**Three ways to import:**\n"
        "- **Individual CSVs** — upload one or more CSV files directly.\n"
        "- **ZIP archive** — upload a ZIP containing one subfolder per test "
        "(each subfolder has the CSV inside; other files are ignored).\n"
        "- **Manage** — rename or delete existing folders and tests below."
    )

    # ── Tab 1: Individual CSVs  |  Tab 2: ZIP ────────────────────────────────
    tab_csv, tab_zip = st.tabs(["📄  Individual CSV files", "📦  ZIP archive"])

    # ── Tab: individual CSVs ──────────────────────────────────────────────────
    with tab_csv:
        exp_name_csv = st.text_input("Experiment name", "Experiment 1",
                                     key="exp_name_csv")
        uploaded = st.file_uploader(
            "Upload CSV files (one per test) — select multiple at once",
            type=['csv'], accept_multiple_files=True, key='csv_uploader')

        if uploaded:
            st.markdown(f"**{len(uploaded)} file(s) ready — review labels:**")
            rows = []
            for i, f in enumerate(uploaded):
                c1, c2, c3 = st.columns([3, 1, 2])
                with c1: st.text(f.name)
                with c2:
                    num = st.number_input("Test #", value=i+1, min_value=0,
                                          step=1, key=f"tnum_{i}",
                                          label_visibility="collapsed")
                with c3:
                    tname = st.text_input(
                        "Label",
                        value=re.sub(r'\.(csv|CSV)$','', f.name),
                        key=f"tname_{i}",
                        label_visibility="collapsed")
                rows.append((f, int(num), tname))

            if st.button("📥  Import all CSV files", type="primary",
                         key="import_csvs"):
                sfid  = db.add_source_folder(exp_name_csv.strip() or "Unnamed")
                ok, fail = 0, []
                for f, num, tname in rows:
                    f.seek(0); raw = f.read()
                    try:
                        load_csv(io.BytesIO(raw))   # validate
                        db.add_test(sfid, tname, num, raw, f.name)
                        ok += 1
                    except Exception as e:
                        fail.append(f"{f.name}: {e}")
                st.success(f"✅  Imported {ok} test(s) into '{exp_name_csv}'.")
                if fail:
                    st.error("Errors:\n" + "\n".join(fail))
                st.rerun()

    # ── Tab: ZIP archive ──────────────────────────────────────────────────────
    with tab_zip:
        st.markdown(
            "Upload a ZIP file. Expected structure:\n"
            "```\n"
            "my_experiment.zip\n"
            "  └── test_01/\n"
            "        ├── data.Stop.csv   ← this is used\n"
            "        └── data.txt        ← ignored\n"
            "  └── test_02/\n"
            "        └── data.Stop.csv\n"
            "```\n"
            "Flat ZIPs (CSVs without subfolders) also work. "
            "Files that aren't CSVs are always ignored."
        )
        exp_name_zip = st.text_input("Experiment name", "Experiment 1",
                                     key="exp_name_zip")
        zip_file = st.file_uploader("Upload ZIP file", type=['zip'],
                                    key='zip_uploader')
        if zip_file is not None:
            if st.button("📥  Import ZIP", type="primary", key="import_zip"):
                with st.spinner("Scanning ZIP and importing CSVs…"):
                    zip_bytes = zip_file.read()
                    count, errors = _import_zip(zip_bytes,
                                                exp_name_zip.strip() or "Unnamed")
                if count:
                    st.success(f"✅  Imported {count} test(s) from ZIP.")
                else:
                    st.error("No valid CSVs found in the ZIP.")
                if errors:
                    with st.expander(f"⚠️ {len(errors)} error(s)"):
                        for e in errors:
                            st.text(e)
                st.rerun()

    # ── Manage existing folders ───────────────────────────────────────────────
    folders = db.get_source_folders()
    if folders:
        st.markdown("---")
        st.subheader("📁  Manage existing folders")
        for sf in folders:
            tests    = db.get_tests_for_folder(sf['id'])
            analyzed = sum(1 for t in tests if t['analyzed'])
            with st.expander(
                f"📁 {_folder_label(sf)}  "
                f"({len(tests)} tests, {analyzed} analyzed)",
                expanded=False):

                col1, col2 = st.columns(2)
                with col1:
                    new_name = st.text_input("Rename", value=_folder_label(sf),
                                             key=f"rf_{sf['id']}")
                    if st.button("✏️ Rename", key=f"ren_{sf['id']}"):
                        db.rename_source_folder(sf['id'], new_name)
                        st.rerun()
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Delete folder + all tests",
                                 key=f"del_{sf['id']}", type="secondary"):
                        db.delete_source_folder(sf['id'])
                        st.rerun()

                for t in tests:
                    tc1, tc2, tc3, tc4 = st.columns([3, 2, 1, 1])
                    with tc1:
                        icon = "✅" if t['analyzed'] else "○"
                        st.text(f"{icon}  {_test_label(t)}")
                    with tc2:
                        tnew = st.text_input("", value=_test_label(t),
                                             key=f"rt_{t['id']}",
                                             label_visibility="collapsed")
                    with tc3:
                        if st.button("✏️", key=f"rtn_{t['id']}"):
                            db.rename_test(t['id'], tnew)
                            st.rerun()
                    with tc4:
                        if st.button("🗑️", key=f"delt_{t['id']}"):
                            db.delete_test(t['id'])
                            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: ANALYSE TESTS
# ─────────────────────────────────────────────────────────────────────────────
elif "Analyse" in page:
    st.title("🔬  Analyse Tests")
    all_tests = db.get_all_tests()
    folders   = db.get_source_folders()

    if not all_tests:
        st.info("No tests imported yet. Go to **Home / Import** first.")
        st.stop()

    # ── Test selector ─────────────────────────────────────────────────────────
    folder_opts = ["All folders"] + [_folder_label(sf) for sf in folders]
    sel_folder  = st.selectbox("Filter by folder", folder_opts, key="at_folder")
    if sel_folder == "All folders":
        visible = all_tests
    else:
        sfid   = next(sf['id'] for sf in folders if _folder_label(sf)==sel_folder)
        visible = [t for t in all_tests if t['source_folder_id']==sfid]

    test_labels = [f"{'✅' if t['analyzed'] else '○'}  {_test_label(t)}"
                   for t in visible]
    sel_label   = st.selectbox("Select test", test_labels, key="at_test")
    t           = visible[test_labels.index(sel_label)]

    x, y = _load_test_data(t)
    if x is None:
        st.error("Could not load CSV data for this test.")
        st.stop()

    # ── Restore / init analysis state ─────────────────────────────────────────
    a    = db.get_analysis(t['id'])
    akey = f"analysis_{t['id']}"
    if akey not in st.session_state:
        st.session_state[akey] = {
            'smoothing': int(a['smoothing_window']) if a else
                         st.session_state.default_smoothing,
            'regions':   json.loads(a['regions_json']) if a else [],
        }
    state = st.session_state[akey]

    left, right = st.columns([1, 2])

    with left:
        st.subheader("Parameters")
        sw = st.slider("Smoothing window", 3, 501, state['smoothing'],
                       step=2, key=f"sw_{t['id']}")
        state['smoothing'] = sw
        ys = smooth_data(y, sw)

        st.markdown("**Analysis Regions**")
        x_min_d, x_max_d = float(x.min()), float(x.max())

        updated = []
        for i, reg in enumerate(state['regions']):
            c = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
            st.markdown(
                f"<span style='color:{c}'>●</span> "
                f"**R{i+1}** {'*(Consistent)*' if reg.get('consistent') else '*(Custom)*'}",
                unsafe_allow_html=True)
            rc1, rc2, rc3 = st.columns([2, 2, 1])
            with rc1:
                xmn = st.number_input("Start mm", value=float(reg['xmin']),
                                      key=f"xmn_{t['id']}_{i}", format="%.3f",
                                      label_visibility="collapsed")
            with rc2:
                xmx = st.number_input("End mm",   value=float(reg['xmax']),
                                      key=f"xmx_{t['id']}_{i}", format="%.3f",
                                      label_visibility="collapsed")
            with rc3:
                if st.button("🗑", key=f"dr_{t['id']}_{i}"):
                    state['regions'].pop(i); st.rerun()
            updated.append({**reg, 'xmin': xmn, 'xmax': xmx, 'color': c})
        state['regions'] = updated

        rc1, rc2 = st.columns(2)
        with rc1:
            nxmn = st.number_input("New start", value=x_min_d,
                                   key=f"nxmn_{t['id']}", format="%.3f")
        with rc2:
            nxmx = st.number_input("New end",   value=x_max_d,
                                   key=f"nxmx_{t['id']}", format="%.3f")
        ra1, ra2 = st.columns(2)
        with ra1:
            if st.button("+ Custom",     key=f"ac_{t['id']}"):
                idx = len(state['regions'])
                state['regions'].append({
                    'xmin': nxmn, 'xmax': nxmx,
                    'color': REGION_COLORS[idx % len(REGION_COLORS)],
                    'type': 'custom', 'consistent': False})
                st.rerun()
        with ra2:
            if st.button("+ Consistent", key=f"acon_{t['id']}"):
                idx = len(state['regions'])
                state['regions'].append({
                    'xmin': nxmn, 'xmax': nxmx,
                    'color': REGION_COLORS[idx % len(REGION_COLORS)],
                    'type': 'consistent', 'consistent': True})
                st.rerun()

        st.markdown("---")
        if st.button("▶  Analyse & Save", type="primary", key=f"run_{t['id']}"):
            ws  = calc_stats(y)
            wss = calc_stats(ys)
            db.save_analysis(t['id'], sw, ws, wss, state['regions'])
            st.success("✅ Analysis saved!")
            st.rerun()

        if a:
            st.markdown("**Results — Whole Range**")
            ws_d  = {k: a.get(f'whole_{k}') or 0 for k in ['mean','max','min','std']}
            wss_d = {k: a.get(f'whole_{k}_smooth') or 0 for k in ['mean','max','min','std']}
            stats_df = pd.DataFrame({
                'Metric': ['Mean','Max','Min','Std'],
                'Raw':    [f"{ws_d[k]:.4f}"  for k in ['mean','max','min','std']],
                'Smooth': [f"{wss_d[k]:.4f}" for k in ['mean','max','min','std']],
            })
            st.dataframe(stats_df, use_container_width=True, hide_index=True)

            regs_saved = json.loads(a.get('regions_json','[]') or '[]')
            if regs_saved:
                st.markdown("**Results — Regions**")
                for i, reg in enumerate(regs_saved):
                    mask = (x >= reg['xmin']) & (x <= reg['xmax'])
                    sr   = calc_stats(y[mask])
                    ss   = calc_stats(smooth_data(y, sw)[mask])
                    c    = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
                    st.markdown(
                        f"<span style='color:{c}'>●</span> "
                        f"**R{i+1}** {reg['xmin']:.2f}→{reg['xmax']:.2f} mm",
                        unsafe_allow_html=True)
                    rd = pd.DataFrame({
                        'Metric': ['Mean','Max','Min','Std'],
                        'Raw':    [f"{sr[k]:.4f}" for k in ['mean','max','min','std']],
                        'Smooth': [f"{ss[k]:.4f}" for k in ['mean','max','min','std']],
                    })
                    st.dataframe(rd, use_container_width=True, hide_index=True)

    with right:
        st.subheader(f"Graph — {_test_label(t)}")
        fig = _draw_test_figure(x, y, ys, state['regions'])
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        buf = io.BytesIO()
        fig2 = _draw_test_figure(x, y, ys, state['regions'])
        fig2.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                     facecolor=st.session_state.graph_bg)
        plt.close(fig2)
        st.download_button("📥 PNG", data=buf.getvalue(),
                           file_name=f"{_test_label(t)}.png", mime="image/png")

    # ── Batch section ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("⚡ Batch Analysis — shared regions")
    bc1, bc2 = st.columns([1, 1])

    with bc1:
        batch_sw = st.slider("Shared smoothing", 3, 501,
                             st.session_state.default_smoothing,
                             step=2, key="batch_sw")
        bkey = "batch_regions"
        if bkey not in st.session_state:
            st.session_state[bkey] = []

        for i, reg in enumerate(st.session_state[bkey]):
            bc_1, bc_2, bc_3 = st.columns([2, 2, 1])
            with bc_1:
                bxmn = st.number_input("From", value=float(reg['xmin']),
                                       key=f"bxmn_{i}", format="%.3f",
                                       label_visibility="collapsed")
            with bc_2:
                bxmx = st.number_input("To",   value=float(reg['xmax']),
                                       key=f"bxmx_{i}", format="%.3f",
                                       label_visibility="collapsed")
            with bc_3:
                if st.button("🗑", key=f"bdr_{i}"):
                    st.session_state[bkey].pop(i); st.rerun()
            st.session_state[bkey][i] = {**reg, 'xmin': bxmn, 'xmax': bxmx}

        if st.button("+ Add shared region", key="addbr"):
            idx = len(st.session_state[bkey])
            st.session_state[bkey].append({
                'xmin': 0.0, 'xmax': 10.0,
                'color': REGION_COLORS[idx % len(REGION_COLORS)],
                'type': 'consistent', 'consistent': True})
            st.rerun()

    with bc2:
        test_opts = [_test_label(t) for t in visible]
        batch_sel = st.multiselect("Tests to batch-analyze",
                                   test_opts, default=test_opts, key="batch_sel")
        batch_tests = [t for t in visible if _test_label(t) in batch_sel]

        if st.button("▶  Run Batch Analysis", type="primary", key="run_batch"):
            n = 0
            for bt in batch_tests:
                bx, by = _load_test_data(bt)
                if bx is None: continue
                bys = smooth_data(by, batch_sw)
                db.save_analysis(bt['id'], batch_sw,
                                 calc_stats(by), calc_stats(bys),
                                 st.session_state[bkey])
                n += 1
            st.success(f"✅ Analyzed {n} test(s).")
            st.rerun()

    # ── XLSX export ───────────────────────────────────────────────────────────
    st.markdown("---")
    analyzed = [t for t in all_tests if t['analyzed']]
    if analyzed:
        st.subheader("📊 Export")
        exp_opts = [_test_label(t) for t in analyzed]
        exp_sel  = st.multiselect("Tests to export", exp_opts, default=exp_opts,
                                  key="exp_sel")
        to_exp   = [t for t in analyzed if _test_label(t) in exp_sel]
        if to_exp:
            st.download_button(
                "📥 Download Excel (.xlsx)",
                data=export_tests_to_xlsx_bytes(to_exp, db),
                file_name="peel_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: SEQUENTIAL ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
elif "Sequential" in page:
    st.title("🔁  Sequential Analysis")

    all_tests = db.get_all_tests()
    folders   = db.get_source_folders()

    # ── State helpers ─────────────────────────────────────────────────────────
    def _seq_reset():
        st.session_state.seq_active   = False
        st.session_state.seq_tests    = []
        st.session_state.seq_idx      = 0
        st.session_state.seq_regions  = []
        st.session_state.seq_done     = False

    def _seq_save_current():
        """Save analysis for the current sequential test."""
        tests = st.session_state.seq_tests
        idx   = st.session_state.seq_idx
        if idx >= len(tests): return
        t     = tests[idx]
        sw    = st.session_state.seq_smoothing
        x, y  = _load_test_data(t)
        if x is None: return
        ys    = smooth_data(y, sw)
        db.save_analysis(
            t['id'], sw,
            calc_stats(y), calc_stats(ys),
            st.session_state.seq_regions)

    # ── Not yet started ───────────────────────────────────────────────────────
    if not st.session_state.seq_active and not st.session_state.seq_done:

        st.info(
            "**Sequential analysis** lets you review each test individually "
            "and define analysis regions one by one.\n\n"
            "1. Choose which tests to analyse.\n"
            "2. Click **Start**.\n"
            "3. For each test: view the graph, set regions, click **Save & Next**.\n"
            "4. When all tests are done, download the Excel results."
        )

        if not all_tests:
            st.warning("Import tests first (Home / Import).")
            st.stop()

        folder_opts = ["All folders"] + [_folder_label(sf) for sf in folders]
        sf_sel = st.selectbox("Filter by folder", folder_opts, key="seq_folder")
        if sf_sel == "All folders":
            visible = all_tests
        else:
            sfid   = next(sf['id'] for sf in folders
                          if _folder_label(sf) == sf_sel)
            visible = [t for t in all_tests if t['source_folder_id'] == sfid]

        test_opts = [_test_label(t) for t in visible]
        sel_tests = st.multiselect("Tests to analyse",
                                   test_opts, default=test_opts,
                                   key="seq_sel_tests")
        chosen = [t for t in visible if _test_label(t) in sel_tests]

        init_sw = st.slider("Initial smoothing window", 3, 501,
                            st.session_state.default_smoothing, step=2,
                            key="seq_init_sw")

        col1, col2 = st.columns(2)
        with col1:
            if chosen and st.button("▶  Start Sequential Analysis",
                                    type="primary", key="seq_start"):
                st.session_state.seq_active   = True
                st.session_state.seq_tests    = chosen
                st.session_state.seq_idx      = 0
                st.session_state.seq_regions  = []
                st.session_state.seq_smoothing = init_sw
                st.session_state.seq_done     = False
                st.rerun()
        with col2:
            if not chosen:
                st.warning("Select at least one test.")

    # ── Done state ────────────────────────────────────────────────────────────
    elif st.session_state.seq_done:
        tests_done = st.session_state.seq_tests
        analyzed   = [t for t in tests_done if db.get_analysis(t['id'])]
        n_done     = len(analyzed)

        st.success(f"✅  Sequential analysis complete — {n_done} / "
                   f"{len(tests_done)} test(s) saved.")

        if analyzed:
            xlsx = export_tests_to_xlsx_bytes(analyzed, db)
            st.download_button(
                "📥  Download Excel results (.xlsx)",
                data=xlsx,
                file_name="sequential_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
                type="primary")
            st.download_button(
                "📤  Download session (.db) with all analyses",
                data=db.to_bytes(),
                file_name="peel_analyzer_session.db",
                mime="application/octet-stream")

        if st.button("🔄  Start a new sequential session", key="seq_restart"):
            _seq_reset()
            st.rerun()

    # ── Active sequential loop ────────────────────────────────────────────────
    else:
        tests     = st.session_state.seq_tests
        idx       = st.session_state.seq_idx
        n_total   = len(tests)

        if idx >= n_total:
            # Reached the end naturally
            st.session_state.seq_active = False
            st.session_state.seq_done   = True
            st.rerun()

        t = tests[idx]

        # ── Progress ──────────────────────────────────────────────────────────
        st.markdown(
            f"### Test {idx + 1} of {n_total} — **{_test_label(t)}**")
        progress = (idx) / n_total
        st.progress(progress,
                    text=f"{idx}/{n_total} saved so far")

        # ── Load data ─────────────────────────────────────────────────────────
        x, y = _load_test_data(t)
        if x is None:
            st.error(f"Cannot load data for {_test_label(t)}. Skipping.")
            st.session_state.seq_idx += 1
            st.session_state.seq_regions = []
            st.rerun()

        x_min_d, x_max_d = float(x.min()), float(x.max())

        # Restore existing analysis if any (pre-fill regions)
        a_existing = db.get_analysis(t['id'])
        if f"seq_init_{t['id']}" not in st.session_state:
            # First time seeing this test in this session
            st.session_state[f"seq_init_{t['id']}"] = True
            if a_existing:
                st.session_state.seq_regions  = json.loads(
                    a_existing.get('regions_json','[]') or '[]')
                st.session_state.seq_smoothing = int(
                    a_existing.get('smoothing_window',
                                   st.session_state.default_smoothing))
            else:
                st.session_state.seq_regions = []

        ys = smooth_data(y, st.session_state.seq_smoothing)

        # ── Two-column layout ─────────────────────────────────────────────────
        left, right = st.columns([1, 2])

        with left:
            # Smoothing
            sw_new = st.slider(
                "Smoothing window", 3, 501,
                st.session_state.seq_smoothing,
                step=2, key=f"seq_sw_{idx}")
            if sw_new != st.session_state.seq_smoothing:
                st.session_state.seq_smoothing = sw_new
                ys = smooth_data(y, sw_new)

            # Regions list
            st.markdown("**Analysis regions:**")
            updated_regs = []
            for i, reg in enumerate(st.session_state.seq_regions):
                c = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
                st.markdown(
                    f"<span style='color:{c}'>●</span> **R{i+1}**",
                    unsafe_allow_html=True)
                sr1, sr2, sr3 = st.columns([2, 2, 1])
                with sr1:
                    rx1 = st.number_input(
                        "Start", value=float(reg['xmin']),
                        key=f"seq_xmn_{idx}_{i}", format="%.3f",
                        label_visibility="collapsed")
                with sr2:
                    rx2 = st.number_input(
                        "End",   value=float(reg['xmax']),
                        key=f"seq_xmx_{idx}_{i}", format="%.3f",
                        label_visibility="collapsed")
                with sr3:
                    if st.button("🗑", key=f"seq_dr_{idx}_{i}"):
                        st.session_state.seq_regions.pop(i)
                        st.rerun()
                updated_regs.append({**reg, 'xmin': rx1, 'xmax': rx2, 'color': c})
            st.session_state.seq_regions = updated_regs

            # Add region
            st.markdown("**Add region:**")
            nr1, nr2 = st.columns(2)
            with nr1:
                new_x0 = st.number_input("From", value=x_min_d,
                                         key=f"seq_nx0_{idx}", format="%.3f")
            with nr2:
                new_x1 = st.number_input("To",   value=x_max_d,
                                         key=f"seq_nx1_{idx}", format="%.3f")
            na1, na2 = st.columns(2)
            with na1:
                if st.button("+ Custom",     key=f"seq_ac_{idx}"):
                    i = len(st.session_state.seq_regions)
                    st.session_state.seq_regions.append({
                        'xmin': new_x0, 'xmax': new_x1,
                        'color': REGION_COLORS[i % len(REGION_COLORS)],
                        'type': 'custom', 'consistent': False})
                    st.rerun()
            with na2:
                if st.button("+ Consistent", key=f"seq_acon_{idx}"):
                    i = len(st.session_state.seq_regions)
                    st.session_state.seq_regions.append({
                        'xmin': new_x0, 'xmax': new_x1,
                        'color': REGION_COLORS[i % len(REGION_COLORS)],
                        'type': 'consistent', 'consistent': True})
                    st.rerun()

            # Whole-range stats preview
            if st.session_state.seq_regions:
                st.markdown("**Preview stats (whole range):**")
                ws  = calc_stats(y)
                wss = calc_stats(ys)
                prev_df = pd.DataFrame({
                    'Metric': ['Mean','Max','Min','Std'],
                    'Raw':    [f"{ws[k]:.4f}"  for k in ['mean','max','min','std']],
                    'Smooth': [f"{wss[k]:.4f}" for k in ['mean','max','min','std']],
                })
                st.dataframe(prev_df, use_container_width=True, hide_index=True)

        with right:
            fig = _draw_test_figure(
                x, y, ys,
                st.session_state.seq_regions,
                title=_test_label(t),
                figsize=(9, 4.5))
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

        # ── Navigation buttons ────────────────────────────────────────────────
        st.markdown("---")
        nav1, nav2, nav3, nav4 = st.columns([2, 2, 2, 2])

        with nav1:
            if idx > 0 and st.button("◀  Previous", key=f"seq_prev_{idx}"):
                # Don't save, just go back
                st.session_state.seq_idx     -= 1
                st.session_state.seq_regions  = []
                # Clear init flag so it re-loads from DB
                prev_t = tests[st.session_state.seq_idx]
                st.session_state.pop(f"seq_init_{prev_t['id']}", None)
                st.rerun()

        with nav2:
            if st.button("⏭  Skip (no save)", key=f"seq_skip_{idx}",
                         type="secondary"):
                st.session_state.seq_idx    += 1
                st.session_state.seq_regions = []
                if idx + 1 < n_total:
                    next_t = tests[idx + 1]
                    st.session_state.pop(f"seq_init_{next_t['id']}", None)
                st.rerun()

        with nav3:
            is_last = (idx == n_total - 1)
            btn_label = "💾  Save & Finish" if is_last else "💾  Save & Next ▶"
            if st.button(btn_label, type="primary", key=f"seq_next_{idx}"):
                _seq_save_current()
                st.session_state.seq_idx    += 1
                st.session_state.seq_regions = []
                if idx + 1 < n_total:
                    next_t = tests[idx + 1]
                    st.session_state.pop(f"seq_init_{next_t['id']}", None)
                if is_last:
                    st.session_state.seq_active = False
                    st.session_state.seq_done   = True
                st.rerun()

        with nav4:
            if st.button("■  Stop & Save current", key=f"seq_stop_{idx}",
                         type="secondary"):
                _seq_save_current()
                st.session_state.seq_active = False
                st.session_state.seq_done   = True
                st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: FIGURE DESIGNER
# ─────────────────────────────────────────────────────────────────────────────
elif "Figure" in page:
    st.title("📊  Figure Designer")
    folders    = db.get_source_folders()
    all_tests  = db.get_all_tests()
    conditions = db.get_conditions()

    # ── Conditions ────────────────────────────────────────────────────────────
    with st.expander("🎨  Manage Conditions", expanded=False):
        for c in conditions:
            cc1, cc2, cc3, cc4 = st.columns([3, 2, 1, 1])
            with cc1:
                cn = st.text_input("", value=c['name'], key=f"cn_{c['id']}",
                                   label_visibility="collapsed")
            with cc2:
                cc = st.color_picker("", value=c['color'], key=f"cc_{c['id']}",
                                     label_visibility="collapsed")
            with cc3:
                if st.button("✏️", key=f"uc_{c['id']}"):
                    db.update_condition(c['id'], cn, cc); st.rerun()
            with cc4:
                if st.button("🗑️", key=f"dc_{c['id']}"):
                    db.delete_condition(c['id']); st.rerun()

        na1, na2, na3 = st.columns([3, 2, 1])
        with na1: nc_n = st.text_input("New name", key="nc_n",
                                       placeholder="e.g. Control")
        with na2: nc_c = st.color_picker("", value="#3b82f6", key="nc_c",
                                         label_visibility="collapsed")
        with na3:
            st.write("")
            if st.button("➕", key="add_c"):
                if nc_n.strip():
                    db.add_condition(nc_n.strip(), nc_c); st.rerun()

    # ── Assignments ───────────────────────────────────────────────────────────
    with st.expander("🔗  Assign Tests to Conditions", expanded=False):
        if not conditions:
            st.info("Create conditions above first.")
        elif not all_tests:
            st.info("Import tests first.")
        else:
            cond_map  = {c['id']: c for c in conditions}
            cond_opts = {c['name']: c['id'] for c in conditions}

            assigned = [t for t in all_tests if t.get('condition_id')]
            if assigned:
                rows = [{'Test': _test_label(t),
                         'Condition': cond_map.get(t['condition_id'],{}).get('name','?')}
                        for t in assigned]
                st.dataframe(pd.DataFrame(rows), use_container_width=True,
                             hide_index=True)
                rm = st.selectbox("Remove assignment",
                                  ["—"] + [_test_label(t) for t in assigned],
                                  key="rm_asgn")
                if rm != "—" and st.button("Remove", key="do_rm"):
                    tid = next(t['id'] for t in assigned if _test_label(t)==rm)
                    db.set_test_condition(tid, None); st.rerun()
            else:
                st.caption("No assignments yet.")

            # Assign
            fol_opts = ["All"] + [_folder_label(sf) for sf in folders]
            fol_sel  = st.selectbox("Filter", fol_opts, key="af_asgn")
            avail    = all_tests if fol_sel == "All" else [
                t for t in all_tests
                if _folder_label(next(
                    (sf for sf in folders if sf['id']==t['source_folder_id']),
                    {})) == fol_sel]

            at_s  = st.selectbox("Test", [_test_label(t) for t in avail],
                                 key="at_asgn")
            cnd_s = st.selectbox("Condition", list(cond_opts.keys()),
                                 key="cnd_asgn")
            if st.button("✅ Assign", type="primary", key="do_asgn"):
                tid = next(t['id'] for t in avail if _test_label(t)==at_s)
                db.set_test_condition(tid, cond_opts[cnd_s]); st.rerun()

    # ── Generate ──────────────────────────────────────────────────────────────
    st.subheader("🖼️  Generate Figure")
    if not conditions:
        st.info("Create conditions and assign tests first.")
    else:
        col_opts, col_fig = st.columns([1, 2])
        with col_opts:
            cond_names = [c['name'] for c in conditions]
            sel_conds  = st.multiselect("Conditions", cond_names,
                                        default=cond_names, key="fig_conds")
            plot_indiv = st.checkbox("Individual curves", value=True)
            plot_mean  = st.checkbox("Mean curve",        value=False)
            sw_fig     = st.slider("Smoothing", 3, 501,
                                   st.session_state.default_smoothing,
                                   step=2, key="fig_sw")
            xlabel = st.text_input("X label", "Displacement (mm)")
            ylabel = st.text_input("Y label",
                                   f"Load ({st.session_state.load_unit})")
            title  = st.text_input("Title", "")
            fw     = st.number_input("Width (cm)",  8.0, 60.0, 20.0, 0.5)
            fh     = st.number_input("Height (cm)", 4.0, 40.0, 12.5, 0.5)
            grid_s = st.selectbox("Grid", ["Major","Major + Minor","None"])
            bg_col = st.color_picker("Background", st.session_state.graph_bg,
                                     key="fig_bg")
            transp = st.checkbox("Transparent background")
            gen    = st.button("▶  Generate", type="primary", key="gen_fig")

            with st.expander("💾 Save / Load config"):
                fn = st.text_input("Config name", key="fig_save_name")
                if st.button("Save", key="save_fig"):
                    if fn.strip():
                        db.save_figure(fn.strip(), {
                            'sel_conds': sel_conds, 'plot_indiv': plot_indiv,
                            'plot_mean': plot_mean, 'sw': sw_fig,
                            'xlabel': xlabel, 'ylabel': ylabel, 'title': title,
                            'fw': fw, 'fh': fh, 'grid': grid_s,
                            'bg': bg_col, 'transp': transp})
                        st.success(f"Saved '{fn.strip()}'"); st.rerun()
                saved = db.get_saved_figures()
                if saved:
                    sf_s = st.selectbox("Load", ["—"]+[f['name'] for f in saved])
                    if sf_s != "—" and st.button("🗑 Delete", key="del_fig"):
                        fid = next(f['id'] for f in saved if f['name']==sf_s)
                        db.delete_figure(fid); st.rerun()

        with col_fig:
            if gen:
                cond_map2 = {c['name']: c for c in conditions}
                sel_objs  = [cond_map2[n] for n in sel_conds if n in cond_map2]
                fig, ax   = plt.subplots(figsize=(fw/2.54, fh/2.54))
                bg        = 'none' if transp else bg_col
                fig.patch.set_facecolor(bg); ax.set_facecolor(bg)
                ax.tick_params(colors='#9ca3af', labelsize=9)
                for sp in ['bottom','left']: ax.spines[sp].set_color('#4b5563')
                for sp in ['top','right']:   ax.spines[sp].set_visible(False)
                ax.set_xlabel(xlabel, color='#9ca3af', fontsize=10)
                ax.set_ylabel(ylabel, color='#9ca3af', fontsize=10)
                if title: ax.set_title(title, color='#f9fafb', fontsize=11,
                                        fontweight='bold')
                if grid_s != "None":
                    ax.grid(True, which='major', color='#374151', lw=0.7, alpha=0.8)
                    if grid_s == "Major + Minor":
                        ax.minorticks_on()
                        ax.grid(True, which='minor', color='#374151',
                                lw=0.3, alpha=0.4)

                handles, plotted = [], False
                for cond in sel_objs:
                    cid  = cond['id']; color = cond['color']
                    cts  = [t for t in all_tests if t.get('condition_id')==cid]
                    axs, ays = [], []
                    for ct in cts:
                        cx, cy = _load_test_data(ct)
                        if cx is None: continue
                        cys = smooth_data(cy, sw_fig)
                        if plot_indiv:
                            ax.plot(cx, cys, color=color, lw=0.9, alpha=0.4)
                            plotted = True
                        axs.append(cx); ays.append(cys)
                    if plot_indiv and axs:
                        handles.append(mpatches.Patch(
                            color=color, label=cond['name'], alpha=0.7))
                    if plot_mean and axs:
                        xmn = max(xi.min() for xi in axs)
                        xmx = min(xi.max() for xi in axs)
                        if xmx > xmn:
                            xc = np.linspace(xmn, xmx, 500)
                            ym = np.mean([np.interp(xc, xi, yi)
                                          for xi, yi in zip(axs, ays)], axis=0)
                            ln, = ax.plot(xc, ym, color=color, lw=2.5,
                                          label=cond['name'])
                            handles = [ln if h.get_label()==cond['name']
                                       else h for h in handles] or [ln]
                            plotted = True
                if handles:
                    ax.legend(handles=handles, facecolor='#1a1d23',
                              labelcolor='#e8eaf0', edgecolor='#374151', fontsize=9)
                fig.tight_layout(pad=1.5)
                if plotted:
                    st.pyplot(fig, use_container_width=True)
                    for fmt, mime, fname in [
                        ('png','image/png','figure.png'),
                        ('pdf','application/pdf','figure.pdf')]:
                        buf = io.BytesIO()
                        fig.savefig(buf, format=fmt, dpi=150,
                                    bbox_inches='tight', facecolor=bg)
                        st.download_button(f"📥 {fmt.upper()}",
                                           data=buf.getvalue(),
                                           file_name=fname, mime=mime)
                else:
                    st.warning("Nothing to plot — assign tests to conditions first.")
                plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
elif "Settings" in page:
    st.title("⚙   Settings")

    st.subheader("Analysis defaults")
    st.session_state.default_smoothing = st.slider(
        "Default smoothing window", 3, 501,
        st.session_state.default_smoothing, step=2)
    st.session_state.load_unit = st.selectbox(
        "Load unit (Y-axis label)", ['N','N/mm','mN/mm'],
        index=['N','N/mm','mN/mm'].index(st.session_state.load_unit))

    st.subheader("Graph colours")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.session_state.graph_bg = st.color_picker(
        "Background", st.session_state.graph_bg)
    with c2: st.session_state.graph_raw_color = st.color_picker(
        "Raw curve", st.session_state.graph_raw_color)
    with c3: st.session_state.graph_smooth_color = st.color_picker(
        "Smooth curve", st.session_state.graph_smooth_color)
    with c4: st.session_state.graph_grid_color = st.color_picker(
        "Grid", st.session_state.graph_grid_color)

    st.subheader("About")
    st.info(
        "**PeelAnalyzer Web** v1.6 — Streamlit edition\n\n"
        "Data lives in memory during the session. "
        "Use **Download session (.db)** in the sidebar before closing."
    )
