import streamlit as st
import pandas as pd
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import (load_file, to_dt, auto_index, parse_summary, calculate_stages,
                   hms_to_min, min_to_hms, hms_to_min_series, hms_to_excel_fraction_series,
                   build_excel_multi)
import io
import xlsxwriter

st.set_page_config(page_title="TAT Analysis", page_icon="📥", layout="wide")

# ── TOP-RIGHT TEMPLATE BUTTONS ────────────────────────────────
_, _c1, _c2 = st.columns([6, 1, 1])
with _c1:
    _ib_h = ["Trip ID","Vehicle Number","Transporter Name","Shift","Gate Entry Type",
             "Supplier Name","Mat. Group","YardIn","GateIn","GrossWeight",
             "TareWeight","GateOut","Net Weight","Remarks"]
    _ib_buf = io.BytesIO()
    _wb = xlsxwriter.Workbook(_ib_buf)
    _ws = _wb.add_worksheet("IB Template")
    _fmt = _wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#1F4E79',
                           'align':'center','valign':'vcenter','border':1,'font_name':'Arial','font_size':10})
    _ws.set_row(0, 28)
    for _i,_h in enumerate(_ib_h):
        _ws.write(0,_i,_h,_fmt)
        _ws.set_column(_i,_i,max(len(_h)+4,16))
    _wb.close(); _ib_buf.seek(0)
    st.download_button("📥 IB Template", data=_ib_buf, file_name="IB_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ib_tmpl_top", use_container_width=True)
with _c2:
    _ob_h = ["Trip ID","Vehicle Number","Transporter Name","Shift","Gate Entry Type",
             "Supplier Name","Mat. Group","YardIn","ParkIn","YardOut","ParkOut",
             "GateIn","TareWeight","LoadingIn","LoadingOut","GrossWeight","GateOut",
             "Unloader Alias","Net Weight","Remarks"]
    _ob_buf = io.BytesIO()
    _wb2 = xlsxwriter.Workbook(_ob_buf)
    _ws2 = _wb2.add_worksheet("OB Template")
    _fmt2 = _wb2.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#1F4E79',
                             'align':'center','valign':'vcenter','border':1,'font_name':'Arial','font_size':10})
    _ws2.set_row(0, 28)
    for _i,_h in enumerate(_ob_h):
        _ws2.write(0,_i,_h,_fmt2)
        _ws2.set_column(_i,_i,max(len(_h)+4,16))
    _wb2.close(); _ob_buf.seek(0)
    st.download_button("📤 OB Template", data=_ob_buf, file_name="OB_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ob_tmpl_top", use_container_width=True)

st.title("📥📤 TAT Analysis")
st.markdown("Calculate TAT for **Inbound** and **Outbound** trips in `HH:MM` format.")
st.markdown("---")

# ── INBOUND / OUTBOUND TOGGLE ─────────────────────────────────
col_btn1, col_btn2, _ = st.columns([1, 1, 4])
with col_btn1:
    if st.button("📥 Inbound", use_container_width=True,
                 type="primary" if st.session_state.get("tat_mode","inbound")=="inbound" else "secondary"):
        st.session_state["tat_mode"] = "inbound"
        for k in ["tat_result","tat_stats","tat_groupby","tat_time_data",
                  "tat_dt_cols","tat_tat_set","tat_total","tat_gateout_col"]:
            st.session_state.pop(k, None)
        st.rerun()
with col_btn2:
    if st.button("📤 Outbound", use_container_width=True,
                 type="primary" if st.session_state.get("tat_mode","inbound")=="outbound" else "secondary"):
        st.session_state["tat_mode"] = "outbound"
        for k in ["tat_result","tat_stats","tat_groupby","tat_time_data",
                  "tat_dt_cols","tat_tat_set","tat_total","tat_gateout_col"]:
            st.session_state.pop(k, None)
        st.rerun()

mode = st.session_state.get("tat_mode", "inbound")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
WEEKDAY_ORDER = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
MONTH_ORDER   = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"]


# ─────────────────────────────────────────────────────────────
# CACHED STAT BUILDERS
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def build_stats_table(result_json, tat_cols_tuple):
    result   = pd.read_json(io.StringIO(result_json))
    tat_cols = list(tat_cols_tuple)
    rows = []
    for col in tat_cols:
        if col not in result.columns: continue
        mins = hms_to_min_series(result[col]).dropna()
        if len(mins) == 0: continue
        rows.append({
            "TAT Stage":  col,
            "Average":    min_to_hms(mins.mean()),
            "Median":     min_to_hms(mins.median()),
            "Min":        min_to_hms(mins.min()),
            "Max":        min_to_hms(mins.max()),
            "Valid Rows": int(mins.count()),
            "Blank Rows": int(len(result) - mins.count()),
        })
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def build_groupby_stats(result_json, tat_cols_tuple, group_col):
    result   = pd.read_json(io.StringIO(result_json))
    tat_cols = list(tat_cols_tuple)
    if not group_col or group_col not in result.columns: return pd.DataFrame()
    min_cols = {col+"_min": hms_to_min_series(result[col])
                for col in tat_cols if col in result.columns}
    tmp  = pd.concat([result[[group_col]], pd.DataFrame(min_cols, index=result.index)], axis=1)
    rows = []
    for col in tat_cols:
        mc = col+"_min"
        if mc not in tmp.columns: continue
        grp = tmp.groupby(group_col)[mc].agg(["mean","median","count"]).reset_index()
        grp.columns = [group_col,"avg","med","cnt"]
        for _, r in grp.iterrows():
            rows.append({"Category":str(r[group_col]),"TAT Stage":col,
                         "Average":min_to_hms(r["avg"]),"Median":min_to_hms(r["med"]),
                         "Trip Count":int(r["cnt"])})
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def build_time_dimension(result_json, tat_cols_tuple, group_col, sort_order_tuple=None):
    result     = pd.read_json(io.StringIO(result_json))
    tat_cols   = list(tat_cols_tuple)
    sort_order = list(sort_order_tuple) if sort_order_tuple else None
    if group_col not in result.columns: return pd.DataFrame()
    min_cols = {col+"_min": hms_to_min_series(result[col])
                for col in tat_cols if col in result.columns}
    tmp  = pd.concat([result[[group_col]], pd.DataFrame(min_cols, index=result.index)], axis=1)
    rows = []
    for col in tat_cols:
        mc = col+"_min"
        if mc not in tmp.columns: continue
        agg = tmp.groupby(group_col)[mc].agg(["mean","median","min","max","count"]).reset_index()
        agg.columns = [group_col,"avg","med","mn","mx","cnt"]
        if sort_order:
            order_map = {k:i for i,k in enumerate(sort_order)}
            agg["_ord"] = agg[group_col].map(lambda x: order_map.get(str(x), 999))
            agg = agg.sort_values("_ord").drop(columns=["_ord"])
        for _, r in agg.iterrows():
            rows.append({group_col:str(r[group_col]),"TAT Stage":col,
                         "Trip Count":int(r["cnt"]),
                         "Average":min_to_hms(r["avg"]),"Median":min_to_hms(r["med"]),
                         "Min":min_to_hms(r["mn"]),"Max":min_to_hms(r["mx"])})
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────
def make_pivot(df, index_col, tat_cols, metric="Average"):
    if df.empty: return pd.DataFrame()
    rows = []
    for key in df[index_col].unique():
        row = {index_col: key}
        sub = df[df[index_col]==key]
        trips = 0
        for _, r in sub.iterrows():
            row[r["TAT Stage"]] = r[metric]
            trips = r["Trip Count"]
        row["Trip Count"] = trips
        rows.append(row)
    return pd.DataFrame(rows)


def show_stat_cards(stats_df):
    if stats_df.empty: return
    cols = st.columns(len(stats_df))
    for i, row in stats_df.iterrows():
        with cols[i]:
            st.markdown(f"""
            <div style='background:#1F4E79;padding:12px;border-radius:10px;
                        text-align:center;margin-bottom:6px'>
                <div style='color:#BDD7EE;font-size:10px;text-transform:uppercase;
                            letter-spacing:1px'>{row['TAT Stage']}</div>
                <div style='color:#00d4ff;font-size:18px;font-weight:700;
                            margin:4px 0'>{row['Average']}</div>
                <div style='color:#BDD7EE;font-size:10px'>Average</div>
            </div>
            <div style='background:#1F4E79;padding:12px;border-radius:10px;text-align:center'>
                <div style='color:#BDD7EE;font-size:10px;text-transform:uppercase;
                            letter-spacing:1px'>{row['TAT Stage']}</div>
                <div style='color:#10b981;font-size:18px;font-weight:700;
                            margin:4px 0'>{row['Median']}</div>
                <div style='color:#BDD7EE;font-size:10px'>Median</div>
            </div>""", unsafe_allow_html=True)


def render_view_table(df, view, extra_cols=None):
    if df.empty: st.info("No data."); return
    base = extra_cols or []
    if   view == "📈 Average": cols = base + ["Average","Trip Count"]
    elif view == "📉 Median":  cols = base + ["Median","Trip Count"]
    else:                       cols = base + ["Average","Median","Min","Max","Trip Count"]
    st.dataframe(df[[c for c in cols if c in df.columns]],
                 use_container_width=True, hide_index=True)


def render_cards(df, label_col, stage_filter=None):
    subset = df[df["TAT Stage"]==stage_filter] if stage_filter else df
    if subset.empty: return
    n = min(len(subset), 7)
    day_cols = st.columns(n)
    for i,(_, row) in enumerate(subset.iterrows()):
        if i >= n: break
        with day_cols[i]:
            st.markdown(f"""
            <div style='background:#1F4E79;padding:8px;border-radius:8px;
                        text-align:center;margin-bottom:4px'>
                <div style='color:#FFFFFF;font-size:9px;font-weight:700'>{row[label_col]}</div>
            </div>
            <div style='background:#2C3E50;padding:7px;border-radius:7px;
                        text-align:center;margin-bottom:3px'>
                <div style='color:#BDD7EE;font-size:9px'>Avg</div>
                <div style='color:#00d4ff;font-size:12px;font-weight:700'>{row['Average']}</div>
            </div>
            <div style='background:#2C3E50;padding:7px;border-radius:7px;
                        text-align:center;margin-bottom:3px'>
                <div style='color:#BDD7EE;font-size:9px'>Median</div>
                <div style='color:#10b981;font-size:12px;font-weight:700'>{row['Median']}</div>
            </div>
            <div style='background:#2C3E50;padding:5px;border-radius:7px;text-align:center'>
                <div style='color:#BDD7EE;font-size:9px'>Trips</div>
                <div style='color:white;font-size:12px;font-weight:700'>{row['Trip Count']}</div>
            </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────
def export_excel(result, stats_df, groupby_df, time_data, dt_col_names, tat_cols_set):
    sheets = {"Full Data": result}
    if not stats_df.empty:    sheets["Overall Stats"]  = stats_df
    if not groupby_df.empty:  sheets["Category Stats"] = groupby_df
    for key, sname in [("date_df","Date Wise"),("dayofwk_df","Day of Week"),
                        ("week_df","Week Wise"),("month_df","Month Wise")]:
        df_td = time_data.get(key, pd.DataFrame())
        if not df_td.empty: sheets[sname] = df_td
    return build_excel_multi(sheets, tat_cols_set=tat_cols_set, dt_col_names=dt_col_names)


# ─────────────────────────────────────────────────────────────
# RENDER ANALYSIS
# ─────────────────────────────────────────────────────────────
def render_analysis():
    result     = st.session_state["tat_result"]
    stats_df   = st.session_state["tat_stats"]
    groupby_df = st.session_state["tat_groupby"]
    time_data  = st.session_state["tat_time_data"]
    dt_cols    = st.session_state["tat_dt_cols"]
    tat_set    = st.session_state["tat_tat_set"]
    total_rows = st.session_state["tat_total"]
    sel_cat    = st.session_state.get("tat_sel_cat","-- None --")
    tat_cols   = list(tat_set)

    # ── OVERALL STATS ─────────────────────────────────────────
    st.markdown("---")
    st.subheader("📊 Overall TAT — Average & Median")
    show_stat_cards(stats_df)
    st.markdown("")
    view = st.radio("Select View", ["📈 Average","📉 Median","📊 Both"],
                    horizontal=True, key="tat_view_radio",
                    index=["📈 Average","📉 Median","📊 Both"].index(
                        st.session_state.get("tat_view","📊 Both")))
    st.session_state["tat_view"] = view
    if not stats_df.empty:
        render_view_table(stats_df, view, extra_cols=["TAT Stage","Valid Rows","Blank Rows"])

    # ── CATEGORY GROUP-BY ──────────────────────────────────────
    if sel_cat != "-- None --" and not groupby_df.empty:
        st.markdown("---")
        st.subheader(f"📋 TAT by {sel_cat}")
        gb_view = st.radio("Category View",["📈 Average","📉 Median","📊 Both"],
                           horizontal=True, key="tat_gb_view")
        for tat_col in tat_cols:
            subset = groupby_df[groupby_df["TAT Stage"]==tat_col].copy()
            if subset.empty: continue
            st.markdown(f"**{tat_col}**")
            render_view_table(subset, gb_view, extra_cols=["Category"])

    # ── TIME DIMENSION TABS ────────────────────────────────────
    st.markdown("---")
    st.subheader("📅 Time Dimension Analysis")
    tab_date, tab_day, tab_week, tab_month = st.tabs([
        "📆 Date Wise","📅 Day of Week","🗓️ Week Wise","🗃️ Month Wise"])

    def render_time_tab(tab, df_key, index_col, label):
        with tab:
            df_t = time_data.get(df_key, pd.DataFrame())
            if df_t.empty:
                st.info(f"No {label} data available.")
                return
            available = df_t["TAT Stage"].unique().tolist()
            sel = st.selectbox("TAT Stage", ["All Stages"]+available, key=f"td_{df_key}_stage")
            dv  = st.radio("View",["📈 Average","📉 Median","📊 Both"],
                           horizontal=True, key=f"td_{df_key}_view")
            st.markdown("")
            if sel == "All Stages":
                st.markdown(f"##### 📊 Pivot — {label} × TAT Stages")
                metric = "Average" if dv=="📈 Average" else "Median"
                st.dataframe(make_pivot(df_t, index_col, tat_cols, metric),
                             use_container_width=True, hide_index=True)
            else:
                sub = df_t[df_t["TAT Stage"]==sel].copy()
                render_view_table(sub, dv, extra_cols=[index_col])
                render_cards(sub, index_col, sel)

    render_time_tab(tab_date,  "date_df",    "GateOut Date",      "Date Wise")
    render_time_tab(tab_day,   "dayofwk_df", "GateOut DayOfWeek", "Day of Week")
    render_time_tab(tab_week,  "week_df",    "GateOut WeekNo",    "Week Wise")
    render_time_tab(tab_month, "month_df",   "GateOut Month",     "Month Wise")

    # ── ROW COUNT ─────────────────────────────────────────────
    st.markdown("---")
    rc1,rc2,rc3 = st.columns(3)
    rc1.metric("Uploaded rows", total_rows)
    rc2.metric("Output rows",   len(result))
    rc3.metric("Match?", "✅ YES" if total_rows==len(result) else "❌ MISMATCH")

    # ── FULL PREVIEW ──────────────────────────────────────────
    st.markdown("---")
    st.subheader(f"👁 Full Data — {len(result)} Rows")
    all_preview = ["Trip ID","Vehicle Number","Transporter Name","Shift",
                   "GateOut Date","GateOut DayOfWeek","GateOut WeekNo","GateOut Month",
                   "YardIn","GateIn","GrossWeight","TareWeight","GateOut",
                   "YI-GI","YI-YO","YO-GI","GI-TW","TW-LI","LI-LO","LO-GW",
                   "GI-GW","GW-TW","TW-GO","GI-GO","GW-GO","GW-LI","LI-LO","LO-TW"]
    prev = list(dict.fromkeys([c for c in all_preview if c in result.columns]))
    st.dataframe(result[prev], use_container_width=True, height=360)

    # ── DOWNLOAD ──────────────────────────────────────────────
    st.markdown("---")
    buf   = export_excel(result, stats_df, groupby_df, time_data, dt_cols, tat_set)
    fname = "Inbound_TAT.xlsx" if mode=="inbound" else "Outbound_TAT.xlsx"
    st.info("✅ Excel: **7 sheets** — Full Data | Overall Stats | Category Stats | Date Wise | Day of Week | Week Wise | Month Wise")
    st.download_button(f"⬇️ Download Excel ({len(result)} rows)",
        data=buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary")


# ─────────────────────────────────────────────────────────────
# SAVE & RERUN
# ─────────────────────────────────────────────────────────────
def save_and_rerun(result, tat_cols_set, dt_col_names, sel_cat, total_rows, gateout_col):
    tat_cols = list(tat_cols_set)

    # Add time dimension columns
    if gateout_col and gateout_col != "-- Not Available --":
        dt_go = to_dt(result[gateout_col])
        if dt_go is not None:
            result["GateOut Date"]      = dt_go.dt.strftime("%d-%m-%Y")
            result["GateOut DayOfWeek"] = dt_go.dt.strftime("%A")
            result["GateOut WeekNo"]    = dt_go.dt.isocalendar().week.astype(str).apply(
                                            lambda w: f"Week {int(w):02d}")
            result["GateOut Month"]     = dt_go.dt.strftime("%B")
            for c in ["GateOut Date","GateOut DayOfWeek","GateOut WeekNo","GateOut Month"]:
                result[c] = result[c].replace("NaT","")

    with st.spinner("⚙️ Computing statistics..."):
        result_json = result.to_json()
        tat_tuple   = tuple(sorted(tat_cols))

        stats   = build_stats_table(result_json, tat_tuple)
        groupby = (build_groupby_stats(result_json, tat_tuple, sel_cat)
                   if sel_cat != "-- None --" else pd.DataFrame())

        time_data = {}
        if "GateOut Date" in result.columns:
            time_data["date_df"] = build_time_dimension(
                result_json, tat_tuple, "GateOut Date")
        if "GateOut DayOfWeek" in result.columns:
            time_data["dayofwk_df"] = build_time_dimension(
                result_json, tat_tuple, "GateOut DayOfWeek",
                sort_order_tuple=tuple(WEEKDAY_ORDER))
        if "GateOut WeekNo" in result.columns:
            all_weeks = tuple(sorted(result["GateOut WeekNo"].dropna().unique().tolist()))
            time_data["week_df"] = build_time_dimension(
                result_json, tat_tuple, "GateOut WeekNo",
                sort_order_tuple=all_weeks)
        if "GateOut Month" in result.columns:
            time_data["month_df"] = build_time_dimension(
                result_json, tat_tuple, "GateOut Month",
                sort_order_tuple=tuple(MONTH_ORDER))

    st.session_state["tat_result"]      = result
    st.session_state["tat_stats"]       = stats
    st.session_state["tat_groupby"]     = groupby
    st.session_state["tat_time_data"]   = time_data
    st.session_state["tat_dt_cols"]     = dt_col_names
    st.session_state["tat_tat_set"]     = tat_cols_set
    st.session_state["tat_total"]       = total_rows
    st.session_state["tat_gateout_col"] = gateout_col
    st.rerun()


# ═══════════════════════════════════════════════════════════════
# INBOUND
# ═══════════════════════════════════════════════════════════════
if mode == "inbound":
    st.subheader("📥 Inbound TAT")
    st.markdown("**Process:** `YardIn → GateIn → GrossWeight → TareWeight → GateOut`")
    with st.expander("📖 Stage Reference"):
        st.markdown("""
        | TAT Column | Formula | Meaning |
        |---|---|---|
        | YI-GI | GateIn − YardIn | Yard to Gate Entry |
        | GI-GW | GrossWeight − GateIn | Gate In to Gross Weighment |
        | GW-TW | TareWeight − GrossWeight | Gross to Tare Weighment |
        | TW-GO | GateOut − TareWeight | Tare to Gate Out |
        | GI-GO | GateOut − GateIn | Total Plant Time |
        """)
    st.markdown("---")

    uploaded = st.file_uploader("📂 Upload Inbound Excel", type=["xlsx","xls"], key="ib_upload")
    if uploaded is None:
        st.info("👆 Upload your inbound Excel file to begin.")
        st.session_state.pop("tat_result", None)
        st.stop()

    df = load_file(uploaded)
    total_rows = len(df)
    st.success(f"✅ Loaded: **{total_rows} rows, {len(df.columns)} columns**")
    with st.expander("🔍 Debug Info"):
        st.write("**Columns:**", df.columns.tolist())
        st.dataframe(df.head(3), use_container_width=True)

    st.markdown("---")
    st.subheader("🔧 Map Your Columns")
    all_cols = ["-- Not Available --"] + df.columns.tolist()
    c1, c2 = st.columns(2)
    with c1:
        col_yardin  = st.selectbox("YardIn",      all_cols, index=auto_index(all_cols,"YardIn"),      key="ib_yi")
        col_gatein  = st.selectbox("GateIn",       all_cols, index=auto_index(all_cols,"GateIn"),      key="ib_gi")
        col_grosswt = st.selectbox("GrossWeight",  all_cols, index=auto_index(all_cols,"GrossWeight"), key="ib_gw")
    with c2:
        col_tarewt  = st.selectbox("TareWeight",   all_cols, index=auto_index(all_cols,"TareWeight"),  key="ib_tw")
        col_gateout = st.selectbox("GateOut",      all_cols, index=auto_index(all_cols,"GateOut"),     key="ib_go")

    st.markdown("---")
    cat_possible = ["Transporter Name","Shift","Mat. Group","Material Group",
                    "Unloader Alias","Vehicle Number","Gate Entry Type","Supplier Name"]
    cat_found = ["-- None --"] + [c for c in cat_possible if c in df.columns]
    sel_cat = st.selectbox("📂 Group-by Category (optional)", cat_found, key="ib_cat")
    st.session_state["tat_sel_cat"] = sel_cat
    st.info("📅 Auto-added: **GateOut Date · DayOfWeek · WeekNo · Month**")
    st.markdown("---")

    if st.button("⚙️ Calculate Inbound TAT", type="primary", use_container_width=True, key="ib_calc"):
        result     = df.copy()
        dt_yardin  = to_dt(result[col_yardin])  if col_yardin  != "-- Not Available --" else None
        dt_gatein  = to_dt(result[col_gatein])  if col_gatein  != "-- Not Available --" else None
        dt_grosswt = to_dt(result[col_grosswt]) if col_grosswt != "-- Not Available --" else None
        dt_tarewt  = to_dt(result[col_tarewt])  if col_tarewt  != "-- Not Available --" else None
        dt_gateout = to_dt(result[col_gateout]) if col_gateout != "-- Not Available --" else None

        st.markdown("#### 📅 DateTime Parse")
        parse_summary([
            ("YardIn",dt_yardin,col_yardin),("GateIn",dt_gatein,col_gatein),
            ("GrossWt",dt_grosswt,col_grosswt),("TareWt",dt_tarewt,col_tarewt),
            ("GateOut",dt_gateout,col_gateout),
        ], st.columns(5))

        st.markdown("#### ⚙️ Calculating Stages")
        stages = [
            ("YI-GI", dt_yardin,  dt_gatein,  "YardIn",      "GateIn"),
            ("GI-GW", dt_gatein,  dt_grosswt, "GateIn",      "GrossWeight"),
            ("GW-TW", dt_grosswt, dt_tarewt,  "GrossWeight", "TareWeight"),
            ("TW-GO", dt_tarewt,  dt_gateout, "TareWeight",  "GateOut"),
            ("GI-GO", dt_gatein,  dt_gateout, "GateIn",      "GateOut"),
        ]
        result, tat_cols_set = calculate_stages(result, stages, st)
        dt_col_names = [c for c in [col_yardin,col_gatein,col_grosswt,col_tarewt,col_gateout]
                        if c != "-- Not Available --"]
        save_and_rerun(result, tat_cols_set, dt_col_names, sel_cat, total_rows, col_gateout)


# ═══════════════════════════════════════════════════════════════
# OUTBOUND
# ═══════════════════════════════════════════════════════════════
else:
    st.subheader("📤 Outbound TAT")
    st.markdown("**Process:** `YardIn → YardOut → GateIn → TareWeight → LoadingIn → LoadingOut → GrossWeight → GateOut`")
    st.markdown("*(ParkIn used as YardIn fallback per row. ParkOut used as YardOut fallback per row.)*")
    with st.expander("📖 Stage Reference"):
        st.markdown("""
        | TAT Column | Formula | Meaning |
        |---|---|---|
        | YI-YO | YardOut − YardIn | Yard dwell time |
        | YI-GI | GateIn − YardIn  | Yard to Gate In |
        | YO-GI | GateIn − YardOut | Yard Out to Gate In |
        | GI-TW | TareWeight − GateIn | Gate In to Tare Weighment |
        | TW-LI | LoadingIn − TareWeight | Tare to Loading Start |
        | LI-LO | LoadingOut − LoadingIn | Loading Duration |
        | LO-GW | GrossWeight − LoadingOut | Loading End to Gross Weighment |
        | GW-GO | GateOut − GrossWeight | Gross to Gate Out |
        | GI-GO | GateOut − GateIn | Total Plant Time |
        """)
    st.markdown("---")

    uploaded = st.file_uploader("📂 Upload Outbound Excel", type=["xlsx","xls"], key="ob_upload")
    if uploaded is None:
        st.info("👆 Upload your outbound Excel file to begin.")
        st.session_state.pop("tat_result", None)
        st.stop()

    df = load_file(uploaded)
    total_rows = len(df)
    st.success(f"✅ Loaded: **{total_rows} rows, {len(df.columns)} columns**")
    with st.expander("🔍 Debug Info"):
        st.write("**Columns:**", df.columns.tolist())
        st.dataframe(df.head(3), use_container_width=True)

    st.markdown("---")
    st.subheader("🔧 Map Your Columns")
    all_cols = ["-- Not Available --"] + df.columns.tolist()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        col_yardin  = st.selectbox("YardIn",  all_cols, index=auto_index(all_cols,"YardIn"),  key="ob_yi")
        col_parkin  = st.selectbox("ParkIn (YardIn fallback)",  all_cols, index=auto_index(all_cols,"ParkIn"),  key="ob_pi")
    with c2:
        col_yardout = st.selectbox("YardOut", all_cols, index=auto_index(all_cols,"YardOut"), key="ob_yo")
        col_parkout = st.selectbox("ParkOut (YardOut fallback)", all_cols, index=auto_index(all_cols,"ParkOut"), key="ob_po")
    with c3:
        col_gatein     = st.selectbox("GateIn",      all_cols, index=auto_index(all_cols,"GateIn"),      key="ob_gi")
        col_tarewt     = st.selectbox("TareWeight",  all_cols, index=auto_index(all_cols,"TareWeight"),  key="ob_tw")
        col_loadingin  = st.selectbox("LoadingIn",   all_cols, index=auto_index(all_cols,"LoadingIn"),   key="ob_li")
    with c4:
        col_loadingout = st.selectbox("LoadingOut",  all_cols, index=auto_index(all_cols,"LoadingOut"),  key="ob_lo")
        col_grosswt    = st.selectbox("GrossWeight", all_cols, index=auto_index(all_cols,"GrossWeight"), key="ob_gw")
        col_gateout    = st.selectbox("GateOut",     all_cols, index=auto_index(all_cols,"GateOut"),     key="ob_go")

    st.info("💡 ParkIn/ParkOut fill blank YardIn/YardOut rows automatically.")
    st.markdown("---")
    cat_possible = ["Transporter Name","Shift","Mat. Group","Material Group",
                    "Unloader Alias","Vehicle Number","Gate Entry Type","Supplier Name"]
    cat_found = ["-- None --"] + [c for c in cat_possible if c in df.columns]
    sel_cat = st.selectbox("📂 Group-by Category (optional)", cat_found, key="ob_cat")
    st.session_state["tat_sel_cat"] = sel_cat
    st.info("📅 Auto-added: **GateOut Date · DayOfWeek · WeekNo · Month**")
    st.markdown("---")

    if st.button("⚙️ Calculate Outbound TAT", type="primary", use_container_width=True, key="ob_calc"):
        result = df.copy()

        dt_yardin     = to_dt(result[col_yardin])     if col_yardin     != "-- Not Available --" else None
        dt_parkin     = to_dt(result[col_parkin])     if col_parkin     != "-- Not Available --" else None
        dt_yardout    = to_dt(result[col_yardout])    if col_yardout    != "-- Not Available --" else None
        dt_parkout    = to_dt(result[col_parkout])    if col_parkout    != "-- Not Available --" else None
        dt_gatein     = to_dt(result[col_gatein])     if col_gatein     != "-- Not Available --" else None
        dt_tarewt     = to_dt(result[col_tarewt])     if col_tarewt     != "-- Not Available --" else None
        dt_loadingin  = to_dt(result[col_loadingin])  if col_loadingin  != "-- Not Available --" else None
        dt_loadingout = to_dt(result[col_loadingout]) if col_loadingout != "-- Not Available --" else None
        dt_grosswt    = to_dt(result[col_grosswt])    if col_grosswt    != "-- Not Available --" else None
        dt_gateout    = to_dt(result[col_gateout])    if col_gateout    != "-- Not Available --" else None

        # ParkIn fallback for YardIn
        if dt_yardin is not None and dt_parkin is not None:
            dt_yardin_eff = dt_yardin.fillna(dt_parkin)
            n = int(dt_yardin.isna().sum() - dt_yardin_eff.isna().sum())
            if n > 0: st.info(f"💡 YardIn: {n} blank rows filled from ParkIn")
        elif dt_yardin is None and dt_parkin is not None:
            dt_yardin_eff = dt_parkin
            st.info("💡 YardIn not mapped — using ParkIn for all rows")
        else:
            dt_yardin_eff = dt_yardin

        # ParkOut fallback for YardOut
        if dt_yardout is not None and dt_parkout is not None:
            dt_yardout_eff = dt_yardout.fillna(dt_parkout)
            n = int(dt_yardout.isna().sum() - dt_yardout_eff.isna().sum())
            if n > 0: st.info(f"💡 YardOut: {n} blank rows filled from ParkOut")
        elif dt_yardout is None and dt_parkout is not None:
            dt_yardout_eff = dt_parkout
            st.info("💡 YardOut not mapped — using ParkOut for all rows")
        else:
            dt_yardout_eff = dt_yardout

        st.markdown("#### 📅 DateTime Parse Summary")
        parse_summary([
            ("YardIn*",  dt_yardin_eff,  f"{col_yardin}+ParkIn"),
            ("YardOut*", dt_yardout_eff, f"{col_yardout}+ParkOut"),
            ("GateIn",   dt_gatein,      col_gatein),
            ("TareWt",   dt_tarewt,      col_tarewt),
            ("GateOut",  dt_gateout,     col_gateout),
        ], st.columns(5))

        st.markdown("#### ⚙️ Calculating Stages")
        stages = []

        # YI-YO: YardIn → YardOut
        if dt_yardin_eff is not None and dt_yardout_eff is not None:
            stages.append(("YI-YO", dt_yardin_eff, dt_yardout_eff, "YardIn*", "YardOut*"))
        # YI-GI: YardIn → GateIn
        if dt_yardin_eff is not None and dt_gatein is not None:
            stages.append(("YI-GI", dt_yardin_eff, dt_gatein, "YardIn*", "GateIn"))
        # YO-GI: YardOut → GateIn
        if dt_yardout_eff is not None and dt_gatein is not None:
            stages.append(("YO-GI", dt_yardout_eff, dt_gatein, "YardOut*", "GateIn"))
        # GI-TW: GateIn → TareWeight
        if dt_gatein is not None and dt_tarewt is not None:
            stages.append(("GI-TW", dt_gatein, dt_tarewt, "GateIn", "TareWeight"))
        # TW-LI: TareWeight → LoadingIn
        if dt_tarewt is not None and dt_loadingin is not None:
            stages.append(("TW-LI", dt_tarewt, dt_loadingin, "TareWeight", "LoadingIn"))
        # LI-LO: LoadingIn → LoadingOut
        if dt_loadingin is not None and dt_loadingout is not None:
            stages.append(("LI-LO", dt_loadingin, dt_loadingout, "LoadingIn", "LoadingOut"))
        # LO-GW: LoadingOut → GrossWeight
        if dt_loadingout is not None and dt_grosswt is not None:
            stages.append(("LO-GW", dt_loadingout, dt_grosswt, "LoadingOut", "GrossWeight"))
        # GW-GO: GrossWeight → GateOut
        if dt_grosswt is not None and dt_gateout is not None:
            stages.append(("GW-GO", dt_grosswt, dt_gateout, "GrossWeight", "GateOut"))
        # GI-GO: GateIn → GateOut (total plant time)
        if dt_gatein is not None and dt_gateout is not None:
            stages.append(("GI-GO", dt_gatein, dt_gateout, "GateIn", "GateOut"))

        result, tat_cols_set = calculate_stages(result, stages, st)
        dt_col_names = [c for c in [col_yardin, col_parkin, col_yardout, col_parkout,
                                     col_gatein, col_tarewt, col_loadingin,
                                     col_loadingout, col_grosswt, col_gateout]
                        if c != "-- Not Available --"]
        save_and_rerun(result, tat_cols_set, dt_col_names, sel_cat, total_rows, col_gateout)


# ─────────────────────────────────────────────────────────────
# RENDER
# ─────────────────────────────────────────────────────────────
if "tat_result" in st.session_state:
    render_analysis()
