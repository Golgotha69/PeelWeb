"""
streamlit_app.py — PeelAnalyzer Web (Streamlit edition)

All analysis logic is identical to the desktop app.
PyQt6 is replaced by Streamlit widgets that run in any browser, including
Chrome on Android.

Workflow
--------
1. Upload your .db session file (optional) OR start fresh.
2. Upload CSV files → they are stored as BLOBs in the in-memory database.
3. Analyse, visualise, assign conditions, build figures.
4. Download the .db session file at any time to resume later.
5. Download results as Excel.

Pages (sidebar radio):
  🏠 Home / Import
  🔬 Analyse Tests
  📊 Figure Designer
  ⚙  Settings
"""

import io
import json
import re
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as ticker
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
#  Session state bootstrap
# ─────────────────────────────────────────────────────────────────────────────
def _init_state():
    if 'db' not in st.session_state:
        st.session_state.db = WebDatabase()
    if 'load_unit' not in st.session_state:
        st.session_state.load_unit = 'N'
    if 'default_smoothing' not in st.session_state:
        st.session_state.default_smoothing = 51
    if 'graph_bg' not in st.session_state:
        st.session_state.graph_bg = '#13151a'
    if 'graph_raw_color' not in st.session_state:
        st.session_state.graph_raw_color = '#3b82f6'
    if 'graph_smooth_color' not in st.session_state:
        st.session_state.graph_smooth_color = '#60a5fa'
    if 'graph_grid_color' not in st.session_state:
        st.session_state.graph_grid_color = '#374151'

_init_state()
db: WebDatabase = st.session_state.db

# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────
REGION_COLORS = ['#f59e0b','#10b981','#ec4899','#8b5cf6','#06b6d4','#ef4444']

def _style_ax(ax, fig, title='', xlabel=None, ylabel=None):
    bg = st.session_state.graph_bg
    ax.set_facecolor(bg)
    fig.patch.set_facecolor(bg)
    ax.tick_params(colors='#9ca3af', labelsize=9)
    for sp in ['bottom','left']:
        ax.spines[sp].set_color('#4b5563')
    for sp in ['top','right']:
        ax.spines[sp].set_visible(False)
    ax.set_xlabel(xlabel or 'Displacement (mm)', color='#9ca3af', fontsize=10)
    ax.set_ylabel(ylabel or f"Load ({st.session_state.load_unit})", color='#9ca3af', fontsize=10)
    ax.grid(True, color=st.session_state.graph_grid_color, linewidth=0.5, alpha=0.6)
    if title:
        ax.set_title(title, color='#f9fafb', fontsize=11, fontweight='bold')

def _load_test_data(t: dict):
    """Return (x, y) numpy arrays for a test, loading from DB blob."""
    raw = db.get_csv_data(t['id'])
    if not raw:
        return None, None
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

# ─────────────────────────────────────────────────────────────────────────────
#  Sidebar — navigation + DB up/download
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔬 PeelAnalyzer Web")
    st.markdown("---")

    page = st.radio("Navigation", [
        "🏠  Home / Import",
        "🔬  Analyse Tests",
        "📊  Figure Designer",
        "⚙   Settings",
    ], label_visibility="collapsed")

    st.markdown("---")
    st.markdown("### 💾 Session Database")

    # Upload existing DB
    db_file = st.file_uploader("Load previous session (.db)", type=['db'],
                               key='db_upload')
    if db_file is not None:
        if st.button("📥 Load this database"):
            raw = db_file.read()
            st.session_state.db = WebDatabase(raw)
            db = st.session_state.db
            st.success("Database loaded!")
            st.rerun()

    # Download current DB
    db_bytes = db.to_bytes()
    st.download_button(
        "📤 Download session (.db)",
        data=db_bytes,
        file_name="peel_analyzer_session.db",
        mime="application/octet-stream",
    )

    st.markdown("---")
    folders = db.get_source_folders()
    st.caption(f"{len(folders)} folder(s)  •  "
               f"{len(db.get_all_tests())} test(s)")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: HOME / IMPORT
# ─────────────────────────────────────────────────────────────────────────────
if "Home" in page:
    st.title("🏠  Home — Import CSV Files")

    st.info(
        "**How to use:**\n"
        "1. Give your experiment a name below.\n"
        "2. Upload one or more CSV files (each file = one test).\n"
        "3. Assign a test number to each, then click **Import**.\n"
        "4. Switch to **Analyse Tests** to view and process them.\n"
        "5. Use **Download session (.db)** in the sidebar to save your work."
    )

    # ── Import new tests ──────────────────────────────────────────────────────
    with st.expander("➕  Import new tests", expanded=True):
        exp_name = st.text_input("Experiment / folder name", "Experiment 1")
        uploaded = st.file_uploader(
            "Upload CSV files (one per test)",
            type=['csv'], accept_multiple_files=True,
            key='csv_uploader')

        if uploaded:
            st.markdown("**Assign test numbers:**")
            rows = []
            for i, f in enumerate(uploaded):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.text(f.name)
                with c2:
                    num = st.number_input("Test #", value=i+1,
                                          min_value=0, step=1,
                                          key=f"tnum_{i}")
                with c3:
                    tname = st.text_input("Label", value=re.sub(r'\.(csv|CSV)$','',f.name),
                                          key=f"tname_{i}")
                rows.append((f, int(num), tname))

            if st.button("📥  Import all", type="primary"):
                sfid = db.add_source_folder(exp_name.strip() or "Unnamed")
                count = 0
                for f, num, tname in rows:
                    f.seek(0)
                    raw = f.read()
                    try:
                        # Validate that load_csv can parse it
                        load_csv(io.BytesIO(raw))
                        db.add_test(sfid, tname, num, raw, f.name)
                        count += 1
                    except Exception as e:
                        st.error(f"❌ {f.name}: {e}")
                st.success(f"✅  Imported {count} test(s) into '{exp_name}'.")
                st.rerun()

    # ── Manage existing folders ───────────────────────────────────────────────
    folders = db.get_source_folders()
    if folders:
        st.markdown("---")
        st.subheader("📁  Existing folders")
        for sf in folders:
            with st.expander(f"📁 {_folder_label(sf)}", expanded=False):
                tests = db.get_tests_for_folder(sf['id'])
                analyzed = sum(1 for t in tests if t['analyzed'])
                st.caption(f"{len(tests)} tests  •  {analyzed} analyzed")

                col1, col2 = st.columns(2)
                with col1:
                    new_name = st.text_input("Rename folder",
                                             value=_folder_label(sf),
                                             key=f"rf_{sf['id']}")
                    if st.button("✏️ Apply rename", key=f"ren_{sf['id']}"):
                        db.rename_source_folder(sf['id'], new_name)
                        st.rerun()
                with col2:
                    if st.button("🗑️ Delete folder", key=f"del_{sf['id']}",
                                 type="secondary"):
                        db.delete_source_folder(sf['id'])
                        st.rerun()

                # List tests
                for t in tests:
                    tc1, tc2, tc3 = st.columns([4, 2, 1])
                    with tc1:
                        icon = "✅" if t['analyzed'] else "○"
                        st.text(f"{icon}  {_test_label(t)}")
                    with tc2:
                        tnew = st.text_input("Rename", value=_test_label(t),
                                             key=f"rt_{t['id']}",
                                             label_visibility="collapsed")
                        if st.button("✏️", key=f"rtn_{t['id']}"):
                            db.rename_test(t['id'], tnew)
                            st.rerun()
                    with tc3:
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
    folder_map  = {sf['id']: sf for sf in folders}
    folder_opts = ["All folders"] + [_folder_label(sf) for sf in folders]
    sel_folder  = st.selectbox("Filter by folder", folder_opts)

    if sel_folder == "All folders":
        visible = all_tests
    else:
        sfid    = next(sf['id'] for sf in folders if _folder_label(sf) == sel_folder)
        visible = [t for t in all_tests if t['source_folder_id'] == sfid]

    test_labels = [f"{'✅' if t['analyzed'] else '○'}  {_test_label(t)}" for t in visible]
    sel_label   = st.selectbox("Select test", test_labels)
    t           = visible[test_labels.index(sel_label)]

    # ── Load data ─────────────────────────────────────────────────────────────
    x, y = _load_test_data(t)
    if x is None:
        st.error("Could not load CSV data for this test.")
        st.stop()

    # ── Restore or init analysis state ────────────────────────────────────────
    a        = db.get_analysis(t['id'])
    akey     = f"analysis_{t['id']}"

    if akey not in st.session_state:
        # Init from saved analysis or defaults
        st.session_state[akey] = {
            'smoothing': int(a['smoothing_window']) if a else st.session_state.default_smoothing,
            'regions':   json.loads(a['regions_json']) if a else [],
        }
    state = st.session_state[akey]

    # ── Two-column layout: controls | graph ───────────────────────────────────
    left, right = st.columns([1, 2])

    with left:
        st.subheader("Parameters")
        sw = st.slider("Smoothing window", 3, 501, state['smoothing'], step=2,
                       key=f"sw_{t['id']}")
        state['smoothing'] = sw
        ys = smooth_data(y, sw)

        # ── Region management ─────────────────────────────────────────────────
        st.markdown("**Analysis Regions**")
        x_min_data = float(x.min())
        x_max_data = float(x.max())

        updated_regions = []
        for i, reg in enumerate(state['regions']):
            color = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
            with st.container():
                st.markdown(
                    f"<span style='color:{color}'>●</span> "
                    f"**R{i+1}** {'(Consistent)' if reg.get('consistent') else '(Custom)'}",
                    unsafe_allow_html=True)
                rc1, rc2, rc3 = st.columns([2, 2, 1])
                with rc1:
                    xmin = st.number_input("Start (mm)", value=float(reg['xmin']),
                                           key=f"xmin_{t['id']}_{i}",
                                           format="%.3f")
                with rc2:
                    xmax = st.number_input("End (mm)", value=float(reg['xmax']),
                                           key=f"xmax_{t['id']}_{i}",
                                           format="%.3f")
                with rc3:
                    if st.button("🗑", key=f"delreg_{t['id']}_{i}"):
                        state['regions'].pop(i)
                        st.rerun()
                updated_regions.append({**reg, 'xmin': xmin, 'xmax': xmax, 'color': color})
        state['regions'] = updated_regions

        # Add region controls
        st.markdown("**Add region:**")
        ac1, ac2 = st.columns(2)
        with ac1:
            new_xmin = st.number_input("From (mm)", value=x_min_data,
                                       key=f"nxmin_{t['id']}", format="%.3f")
        with ac2:
            new_xmax = st.number_input("To (mm)", value=x_max_data,
                                       key=f"nxmax_{t['id']}", format="%.3f")
        ac3, ac4 = st.columns(2)
        with ac3:
            if st.button("+ Custom", key=f"addcus_{t['id']}"):
                idx = len(state['regions'])
                state['regions'].append({
                    'xmin': new_xmin, 'xmax': new_xmax,
                    'color': REGION_COLORS[idx % len(REGION_COLORS)],
                    'type': 'custom', 'consistent': False})
                st.rerun()
        with ac4:
            if st.button("+ Consistent", key=f"addcon_{t['id']}"):
                idx = len(state['regions'])
                state['regions'].append({
                    'xmin': new_xmin, 'xmax': new_xmax,
                    'color': REGION_COLORS[idx % len(REGION_COLORS)],
                    'type': 'consistent', 'consistent': True})
                st.rerun()

        st.markdown("---")
        if st.button("▶  Analyse This Test", type="primary", key=f"run_{t['id']}"):
            ws  = calc_stats(y)
            wss = calc_stats(ys)
            db.save_analysis(t['id'], sw, ws, wss, state['regions'])
            st.success("✅ Analysis saved!")
            st.rerun()

        # ── Whole-range stats ──────────────────────────────────────────────────
        if a:
            st.markdown("**Results — Whole Range**")
            ws_d  = {k: a.get(f'whole_{k}') or 0 for k in ['mean','max','min','std']}
            wss_d = {k: a.get(f'whole_{k}_smooth') or 0 for k in ['mean','max','min','std']}
            stats_df = pd.DataFrame({
                'Metric': ['Mean','Max','Min','Std Dev'],
                'Raw':    [f"{ws_d[k]:.4f}" for k in ['mean','max','min','std']],
                'Smooth': [f"{wss_d[k]:.4f}" for k in ['mean','max','min','std']],
            })
            st.dataframe(stats_df, use_container_width=True, hide_index=True)

            # Per-region stats
            regions_saved = json.loads(a.get('regions_json','[]') or '[]')
            if regions_saved:
                st.markdown("**Results — Regions**")
                for i, reg in enumerate(regions_saved):
                    mask = (x >= reg['xmin']) & (x <= reg['xmax'])
                    yr   = y[mask]; ysr = ys[mask]
                    sr   = calc_stats(yr); ss = calc_stats(ysr)
                    color = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
                    st.markdown(
                        f"<span style='color:{color}'>●</span> "
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
        fig, ax = plt.subplots(figsize=(9, 5))
        _style_ax(ax, fig)

        ax.plot(x, y, color=st.session_state.graph_raw_color,
                lw=0.8, alpha=0.4, label='Raw')
        ax.plot(x, ys, color=st.session_state.graph_smooth_color,
                lw=1.8, label='Smoothed')

        for i, reg in enumerate(state['regions']):
            color = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
            ax.axvspan(reg['xmin'], reg['xmax'], alpha=0.15, color=color, zorder=3)
            for xv in [reg['xmin'], reg['xmax']]:
                ax.axvline(xv, color=color, lw=1, ls='--', alpha=0.7)

        ax.legend(fontsize=8, facecolor='#1a1d23', labelcolor='#e8eaf0',
                  edgecolor='#374151', framealpha=0.9)
        fig.tight_layout(pad=1.5)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        # Download graph as PNG
        buf = io.BytesIO()
        fig2, ax2 = plt.subplots(figsize=(9, 5))
        _style_ax(ax2, fig2)
        ax2.plot(x, y, color=st.session_state.graph_raw_color, lw=0.8, alpha=0.4, label='Raw')
        ax2.plot(x, ys, color=st.session_state.graph_smooth_color, lw=1.8, label='Smoothed')
        for i, reg in enumerate(state['regions']):
            color = reg.get('color', REGION_COLORS[i % len(REGION_COLORS)])
            ax2.axvspan(reg['xmin'], reg['xmax'], alpha=0.15, color=color, zorder=3)
        ax2.legend(fontsize=8, facecolor='#1a1d23', labelcolor='#e8eaf0',
                   edgecolor='#374151', framealpha=0.9)
        fig2.tight_layout(pad=1.5)
        fig2.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                     facecolor=st.session_state.graph_bg)
        plt.close(fig2)
        st.download_button("📥 Download graph (PNG)",
                           data=buf.getvalue(),
                           file_name=f"{_test_label(t)}.png",
                           mime="image/png")

    # ── Batch operations ──────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("⚡ Batch Operations")

    bcol1, bcol2 = st.columns(2)
    with bcol1:
        st.markdown("**Shared regions for batch analysis**")
        st.caption("Define regions here, apply to selected tests with shared smoothing.")
        batch_sw = st.slider("Shared smoothing window", 3, 501,
                             st.session_state.default_smoothing, step=2,
                             key="batch_sw")
        bkey = "batch_regions"
        if bkey not in st.session_state:
            st.session_state[bkey] = []

        for i, reg in enumerate(st.session_state[bkey]):
            bc1, bc2, bc3 = st.columns([2,2,1])
            with bc1: bxmin = st.number_input("From", value=float(reg['xmin']),
                                               key=f"bxmin_{i}", format="%.3f")
            with bc2: bxmax = st.number_input("To",   value=float(reg['xmax']),
                                               key=f"bxmax_{i}", format="%.3f")
            with bc3:
                if st.button("🗑", key=f"bdelreg_{i}"):
                    st.session_state[bkey].pop(i); st.rerun()
            st.session_state[bkey][i] = {**reg, 'xmin': bxmin, 'xmax': bxmax}

        if st.button("+ Add shared region", key="addbatch"):
            idx = len(st.session_state[bkey])
            st.session_state[bkey].append({
                'xmin': 0.0, 'xmax': 10.0,
                'color': REGION_COLORS[idx % len(REGION_COLORS)],
                'type': 'consistent', 'consistent': True})
            st.rerun()

    with bcol2:
        st.markdown("**Select tests to batch-analyze**")
        test_options = [_test_label(t) for t in visible]
        batch_sel = st.multiselect("Tests", test_options, default=test_options)
        batch_tests = [t for t in visible if _test_label(t) in batch_sel]

        if st.button("▶  Run Batch Analysis", type="primary", key="run_batch"):
            count = 0
            for bt in batch_tests:
                bx, by = _load_test_data(bt)
                if bx is None: continue
                bys = smooth_data(by, batch_sw)
                ws  = calc_stats(by)
                wss = calc_stats(bys)
                db.save_analysis(bt['id'], batch_sw, ws, wss,
                                 st.session_state[bkey])
                count += 1
            st.success(f"✅ Analyzed {count} test(s).")
            st.rerun()

    # ── XLSX export ───────────────────────────────────────────────────────────
    st.markdown("---")
    analyzed_tests = [t for t in all_tests if t['analyzed']]
    if analyzed_tests:
        st.subheader("📊 Export Results")
        export_sel = st.multiselect(
            "Tests to export (analyzed only)",
            [_test_label(t) for t in analyzed_tests],
            default=[_test_label(t) for t in analyzed_tests])
        to_export = [t for t in analyzed_tests if _test_label(t) in export_sel]
        if to_export:
            xlsx_bytes = export_tests_to_xlsx_bytes(to_export, db)
            st.download_button("📥 Download Excel (.xlsx)",
                               data=xlsx_bytes,
                               file_name="peel_analyzer_results.xlsx",
                               mime="application/vnd.openxmlformats-officedocument"
                                    ".spreadsheetml.sheet")

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: FIGURE DESIGNER
# ─────────────────────────────────────────────────────────────────────────────
elif "Figure" in page:
    st.title("📊  Figure Designer")

    folders    = db.get_source_folders()
    all_tests  = db.get_all_tests()
    conditions = db.get_conditions()

    # ── Conditions management ──────────────────────────────────────────────────
    with st.expander("🎨  Manage Conditions", expanded=False):
        if conditions:
            st.markdown("**Existing conditions** (double-click name to rename):")
            for c in conditions:
                cc1, cc2, cc3, cc4 = st.columns([3, 2, 1, 1])
                with cc1:
                    new_cname = st.text_input("Name", value=c['name'],
                                              key=f"cname_{c['id']}",
                                              label_visibility="collapsed")
                with cc2:
                    new_ccolor = st.color_picker("Color", value=c['color'],
                                                 key=f"ccolor_{c['id']}",
                                                 label_visibility="collapsed")
                with cc3:
                    if st.button("✏️", key=f"updcond_{c['id']}"):
                        db.update_condition(c['id'], new_cname, new_ccolor)
                        st.rerun()
                with cc4:
                    if st.button("🗑️", key=f"delcond_{c['id']}"):
                        db.delete_condition(c['id'])
                        st.rerun()

        st.markdown("**Add condition:**")
        na1, na2, na3 = st.columns([3, 2, 1])
        with na1: nc_name  = st.text_input("Name", key="nc_name",
                                            placeholder="e.g. Control")
        with na2: nc_color = st.color_picker("Color", value="#3b82f6",
                                              key="nc_color")
        with na3:
            st.write("")
            if st.button("➕ Add", key="add_cond"):
                if nc_name.strip():
                    db.add_condition(nc_name.strip(), nc_color)
                    st.rerun()

    # ── Test assignment ─────────────────────────────────────────────────────────
    with st.expander("🔗  Assign Tests to Conditions", expanded=False):
        if not conditions:
            st.info("Create at least one condition above first.")
        elif not all_tests:
            st.info("Import tests first.")
        else:
            cond_map  = {c['id']: c for c in conditions}
            cond_opts = {c['name']: c['id'] for c in conditions}

            st.markdown("**Current assignments:**")
            assigned = [t for t in all_tests if t.get('condition_id')]
            if assigned:
                rows = []
                for t in assigned:
                    cond = cond_map.get(t['condition_id'], {})
                    rows.append({'Test': _test_label(t), 'Condition': cond.get('name','?'),
                                 'Color': cond.get('color','')})
                adf = pd.DataFrame(rows)
                st.dataframe(adf[['Test','Condition']], use_container_width=True,
                             hide_index=True)

                # Remove assignment
                rm_sel = st.selectbox("Remove assignment for",
                                      ["— select —"] + [_test_label(t) for t in assigned],
                                      key="rm_assign")
                if rm_sel != "— select —" and st.button("Remove", key="do_rm"):
                    tid = next(t['id'] for t in assigned if _test_label(t) == rm_sel)
                    db.set_test_condition(tid, None)
                    st.rerun()
            else:
                st.caption("No tests assigned yet.")

            st.markdown("**Assign:**")
            folder_opts2 = ["All folders"] + [_folder_label(sf) for sf in folders]
            af_sel = st.selectbox("Filter folder", folder_opts2, key="af_assign")
            if af_sel == "All folders":
                avail_tests = all_tests
            else:
                sfid2 = next(sf['id'] for sf in folders if _folder_label(sf) == af_sel)
                avail_tests = [t for t in all_tests if t['source_folder_id'] == sfid2]

            at_sel  = st.selectbox("Test", [_test_label(t) for t in avail_tests],
                                   key="at_assign")
            cnd_sel = st.selectbox("Condition", list(cond_opts.keys()), key="cnd_assign")
            if st.button("✅ Assign", key="do_assign", type="primary"):
                tid  = next(t['id'] for t in avail_tests if _test_label(t) == at_sel)
                cid  = cond_opts[cnd_sel]
                db.set_test_condition(tid, cid)
                st.rerun()

    # ── Plot options ───────────────────────────────────────────────────────────
    st.subheader("🖼️  Generate Figure")
    if not conditions:
        st.info("Create conditions and assign tests first.")
    else:
        col_opts, col_fig = st.columns([1, 2])

        with col_opts:
            cond_names = [c['name'] for c in conditions]
            sel_conds  = st.multiselect("Conditions to include",
                                        cond_names, default=cond_names,
                                        key="fig_conds")

            plot_indiv   = st.checkbox("Individual curves", value=True)
            plot_mean    = st.checkbox("Mean curve", value=False)

            sw_fig = st.slider("Smoothing", 3, 501,
                               st.session_state.default_smoothing, step=2,
                               key="fig_sw")

            xlabel = st.text_input("X label", "Displacement (mm)")
            ylabel = st.text_input("Y label", f"Load ({st.session_state.load_unit})")
            title  = st.text_input("Title", "")

            fw = st.number_input("Width (cm)", 8.0, 60.0, 20.0, 0.5)
            fh = st.number_input("Height (cm)", 4.0, 40.0, 12.5, 0.5)

            grid_style = st.selectbox("Grid", ["Major", "Major + Minor", "None"])

            bg_color = st.color_picker("Background", st.session_state.graph_bg,
                                       key="fig_bg")
            transp   = st.checkbox("Transparent background")

            gen_fig = st.button("▶  Generate", type="primary", key="gen_fig")

            # Save figure config
            with st.expander("💾 Save / Load figure config"):
                fig_name = st.text_input("Config name", key="fig_save_name")
                if st.button("Save config", key="save_fig_cfg"):
                    if fig_name.strip():
                        cfg = {
                            'sel_conds': sel_conds, 'plot_indiv': plot_indiv,
                            'plot_mean': plot_mean, 'sw': sw_fig,
                            'xlabel': xlabel, 'ylabel': ylabel, 'title': title,
                            'fw': fw, 'fh': fh, 'grid': grid_style,
                            'bg': bg_color, 'transp': transp,
                        }
                        db.save_figure(fig_name.strip(), cfg)
                        st.success(f"Saved '{fig_name.strip()}'")
                        st.rerun()

                saved_figs = db.get_saved_figures()
                if saved_figs:
                    sf_sel = st.selectbox("Load saved config",
                                          ["—"] + [f['name'] for f in saved_figs])
                    if sf_sel != "—" and st.button("Load", key="load_fig_cfg"):
                        fig_row = next(f for f in saved_figs if f['name'] == sf_sel)
                        cfg = json.loads(fig_row['config_json'])
                        st.info(f"Loaded '{sf_sel}'. Re-generate the figure.")
                    if sf_sel != "—" and st.button("🗑 Delete", key="del_fig_cfg"):
                        fig_row = next(f for f in saved_figs if f['name'] == sf_sel)
                        db.delete_figure(fig_row['id'])
                        st.rerun()

        with col_fig:
            if gen_fig:
                cond_map2 = {c['name']: c for c in conditions}
                sel_cond_objs = [cond_map2[n] for n in sel_conds if n in cond_map2]

                fig, ax = plt.subplots(figsize=(fw/2.54, fh/2.54))
                bg = 'none' if transp else bg_color
                fig.patch.set_facecolor(bg)
                ax.set_facecolor(bg)
                ax.tick_params(colors='#9ca3af', labelsize=9)
                for sp in ['bottom','left']: ax.spines[sp].set_color('#4b5563')
                for sp in ['top','right']:   ax.spines[sp].set_visible(False)
                ax.set_xlabel(xlabel, color='#9ca3af', fontsize=10)
                ax.set_ylabel(ylabel, color='#9ca3af', fontsize=10)
                if title: ax.set_title(title, color='#f9fafb', fontsize=11, fontweight='bold')

                # Grid
                if grid_style != "None":
                    ax.grid(True, which='major', color='#374151', lw=0.7, alpha=0.8)
                    if grid_style == "Major + Minor":
                        ax.minorticks_on()
                        ax.grid(True, which='minor', color='#374151', lw=0.3, alpha=0.4)

                legend_handles = []
                plotted = False

                for cond in sel_cond_objs:
                    cid        = cond['id']
                    color      = cond['color']
                    cond_tests = [t for t in all_tests if t.get('condition_id') == cid]
                    all_xs, all_ys = [], []

                    for ct in cond_tests:
                        cx, cy = _load_test_data(ct)
                        if cx is None: continue
                        cys = smooth_data(cy, sw_fig)
                        if plot_indiv:
                            ax.plot(cx, cys, color=color, lw=0.9, alpha=0.4)
                            plotted = True
                        all_xs.append(cx); all_ys.append(cys)

                    drew_indiv = plot_indiv and len(all_xs) > 0
                    if drew_indiv:
                        legend_handles.append(
                            mpatches.Patch(color=color, label=cond['name'], alpha=0.7))
                        plotted = True

                    if plot_mean and all_xs:
                        xmin_c = max(xi.min() for xi in all_xs)
                        xmax_c = min(xi.max() for xi in all_xs)
                        if xmax_c > xmin_c:
                            xc   = np.linspace(xmin_c, xmax_c, 500)
                            ym   = np.mean([np.interp(xc, xi, yi)
                                            for xi, yi in zip(all_xs, all_ys)], axis=0)
                            line, = ax.plot(xc, ym, color=color, lw=2.5,
                                            label=cond['name'])
                            if drew_indiv:
                                legend_handles[-1] = line
                            else:
                                legend_handles.append(line)
                            plotted = True

                if legend_handles:
                    ax.legend(handles=legend_handles, facecolor='#1a1d23',
                              labelcolor='#e8eaf0', edgecolor='#374151', fontsize=9)
                fig.tight_layout(pad=1.5)

                if plotted:
                    st.pyplot(fig, use_container_width=True)
                    buf = io.BytesIO()
                    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                                facecolor=bg)
                    st.download_button("📥 Download figure (PNG)",
                                       data=buf.getvalue(),
                                       file_name="figure.png",
                                       mime="image/png")
                    buf2 = io.BytesIO()
                    fig.savefig(buf2, format='pdf', bbox_inches='tight',
                                facecolor=bg)
                    st.download_button("📥 Download figure (PDF)",
                                       data=buf2.getvalue(),
                                       file_name="figure.pdf",
                                       mime="application/pdf")
                else:
                    st.warning("Nothing to plot. Assign tests to the selected conditions first.")
                plt.close(fig)

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE: SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
elif "Settings" in page:
    st.title("⚙   Settings")

    st.subheader("Analysis defaults")
    new_sw = st.slider("Default smoothing window",
                       3, 501, st.session_state.default_smoothing, step=2)
    st.session_state.default_smoothing = new_sw

    new_unit = st.selectbox("Load unit (Y-axis label)",
                             ['N', 'N/mm', 'mN/mm'],
                             index=['N','N/mm','mN/mm'].index(st.session_state.load_unit))
    st.session_state.load_unit = new_unit

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
        "**PeelAnalyzer Web** — Streamlit edition\n\n"
        "This is a browser-based version of PeelAnalyzer. "
        "Data is stored in an in-memory SQLite database during your session. "
        "**Download your session (.db) in the sidebar** before closing the tab "
        "to save your work and re-upload it next time."
    )
    st.markdown("**Desktop app requirements:** PyQt6, matplotlib, numpy, scipy, pandas, openpyxl")
    st.markdown("**Web app requirements:** streamlit, matplotlib, numpy, scipy, pandas, openpyxl")
