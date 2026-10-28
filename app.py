import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# --- 1. ENTERPRISE PAGE CONFIG ---
st.set_page_config(
    page_title="Pendency Pulse 360",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. MODERN ENTERPRISE CSS ---
st.markdown("""
<style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 98%; }
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: white; padding: 18px 24px; border-radius: 14px;
        display: flex; justify-content: space-between; align-items: center;
        margin-bottom: 18px; box-shadow: 0 4px 15px rgba(0,0,0,0.12);
        border-left: 6px solid #f59e0b;
    }
    .header-title { font-size: 26px; font-weight: 800; color: #ffffff; margin: 0; }
    .header-sub { font-size: 13px; color: #94a3b8; margin-top: 3px; }
    .jagan-badge {
        background: rgba(245, 158, 11, 0.15); border: 1px solid #f59e0b;
        color: #fbbf24; padding: 6px 14px; border-radius: 20px;
        font-size: 12px; font-weight: 600;
    }
    .kpi-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 14px 16px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        text-align: center; margin-bottom: 10px;
    }
    .kpi-title { font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase; margin-bottom: 4px; }
    .kpi-val { font-size: 24px; font-weight: 800; color: #0f172a; }
    .kpi-crit { border-bottom: 4px solid #ef4444; }
    .kpi-warn { border-bottom: 4px solid #f59e0b; }
    .kpi-good { border-bottom: 4px solid #10b981; }
    .kpi-info { border-bottom: 4px solid #3b82f6; }
    .kpi-purp { border-bottom: 4px solid #8b5cf6; }
    .mini-card {
        background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 10px;
        padding: 10px 12px; text-align: center;
    }
    .mini-title { font-size: 11px; font-weight: 700; color: #475569; }
    .mini-val { font-size: 19px; font-weight: 800; color: #1e293b; }
    .drill-box {
        background: #eff6ff; border-left: 5px solid #2563eb;
        padding: 12px 16px; border-radius: 8px; margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. SMART AUTO-LOGIN ---
if st.query_params.get("auth") != "success":
    st.markdown("""
    <div style='max-width: 420px; margin: 80px auto; padding: 30px; background: #0f172a; color: white; border-radius: 16px; text-align: center;'>
        <h2 style='color: #f59e0b; margin-bottom: 5px;'>⚡ Pendency Pulse 360</h2>
        <p style='color: #94a3b8; font-size: 13px;'>Secure Operations Intelligence Portal</p>
    </div>
    """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        pwd = st.text_input("🔑 Enter Hub Access Password:", type="password")
        if st.button("Unlock Portal 🚀", use_container_width=True):
            if pwd == st.secrets.get("APP_PASSWORD", "SuratHub@2026"):
                st.query_params["auth"] = "success"
                st.rerun()
            else:
                st.error("❌ Invalid Password! Access Denied.")
    st.stop()

# --- 4. TOP MODERN HEADER ---
st.markdown("""
<div class="main-header">
    <div>
        <div class="header-title">⚡ PENDENCY PULSE 360</div>
        <div class="header-sub">Live Bag Connection & Processing Command Center • SRT Hub Operations</div>
    </div>
    <div class="jagan-badge">⚡ Crafted by Jagan</div>
</div>
""", unsafe_allow_html=True)

# --- 5. SIDEBAR: DATA UPLOAD & PERSISTENCE ---
st.sidebar.markdown("### 📂 Control Center")
st.sidebar.caption("Ek baar upload karein, data server par save rahega.")

map_up = st.sidebar.file_uploader("1. Upload Mapping (Expriment 2.xlsx)", type=['xlsx'])
in_up = st.sidebar.file_uploader("2. Upload InScan CSV", type=['csv'])
cl_up = st.sidebar.file_uploader("3. Upload BagClosed CSV", type=['csv'])

if map_up:
    with open("map.xlsx", "wb") as f:
        f.write(map_up.getbuffer())
if in_up:
    with open("in.csv", "wb") as f:
        f.write(in_up.getbuffer())
if cl_up:
    with open("cl.csv", "wb") as f:
        f.write(cl_up.getbuffer())

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    return output.getvalue()

def find_col(df, candidates, default_val="-"):
    col_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in col_map:
            return df[col_map[cand.lower()]].fillna(default_val).astype(str)
    return pd.Series([default_val] * len(df), index=df.index)

# --- 6. CORE DATA ENGINE ---
if os.path.exists("map.xlsx") and os.path.exists("in.csv") and os.path.exists("cl.csv"):

    @st.cache_data(ttl=60)
    def load_mappings(mtime):
        try:
            xls = pd.ExcelFile("map.xlsx")
            dest_df = pd.read_excel(xls, "GroupName").fillna("")
            lane_df = pd.read_excel(xls, "lANE").fillna("")
            k1 = dest_df['ShipmentDestination'].astype(str).str.strip().str.upper()
            v1 = dest_df['GroupName'].astype(str).str.strip()
            k2 = lane_df['Groups'].astype(str).str.strip().str.upper()
            v2 = lane_df['lANE'].astype(str).str.strip()
            return dict(zip(k1, v1)), dict(zip(k2, v2))
        except Exception:
            return {}, {}

    dest_map, lane_map = load_mappings(os.path.getmtime("map.xlsx"))
    in_df_raw = pd.read_csv("in.csv")
    cl_df_raw = pd.read_csv("cl.csv")

    AGE_BUCKETS = ['0-1 hr', '1-3 hrs', '3-6 hrs', '6-12 hrs', '12-24 hrs', '>24 hrs']

    def standardize_df(df, source_label):
        df = df.copy()
        df['Source'] = source_label
        df['bagno'] = find_col(df, ['bagno', 'bag_no', 'bagnumber', 'bag id'], 'UNKNOWN').str.strip()
        
        c_ship = find_col(df, ['totalshipments', 'shipments', 'shipmentcount'], 0)
        df['totalshipments'] = pd.to_numeric(c_ship, errors='coerce').fillna(0)
        
        c_wt = find_col(df, ['physicalweight', 'weight', 'bagweight'], 0)
        df['physicalweight'] = pd.to_numeric(c_wt, errors='coerce').fillna(0)
        
        c_hrs = find_col(df, ['pendinghrs', 'pending_hrs', 'ageing', 'pendinghours'], 0)
        df['pendinghrs'] = pd.to_numeric(c_hrs, errors='coerce').fillna(0)
        
        raw_dest = find_col(df, ['destinationhubname', 'destination', 'desthub'], 'UNKNOWN')
        df['Clean_Dest'] = raw_dest.apply(lambda x: str(x).split('(')[0].strip().upper())
        df['Current_Hub'] = find_col(df, ['currenthubname', 'currenthub', 'hubname', 'origin'], 'SRT/PC2')
        df['Created_Time'] = find_col(df, ['bagcreateddate', 'createddate', 'bagcreatedtime', 'created_at'], '-')
        df['Created_By'] = find_col(df, ['bagcreatedby', 'createdby', 'created_by', 'user'], '-')
        df['Last_Modified_By'] = find_col(df, ['lastmodifiedby', 'modifiedby', 'last_modified_by'], '-')
        df['Bag_Type'] = find_col(df, ['bagtype', 'bag_type', 'type'], 'Normal')
        df['Bag_Mode'] = find_col(df, ['bagmode', 'mode', 'servicemode', 'transitmode'], 'Surface')
        
        df['Mapped_Group'] = df['Clean_Dest'].apply(lambda x: dest_map.get(x, "UNMAPPED"))
        df['Mapped_Lane'] = df['Mapped_Group'].apply(
            lambda x: lane_map.get(x.upper(), "UNMAPPED") if x != "UNMAPPED" else "UNMAPPED"
        )
        
        def classify_row(r):
            cond1 = r['Clean_Dest'].startswith("SRT")
            cond2 = r['Mapped_Group'].upper() == "SURAT CITY"
            cond3 = r['Mapped_Lane'].upper() == "LOCAL"
            return "Processing" if (cond1 or cond2 or cond3) else "Connection"
            
        df['Classification'] = df.apply(classify_row, axis=1)
        df['Age_Bucket'] = pd.cut(
            df['pendinghrs'],
            bins=[-1, 1, 3, 6, 12, 24, 999999],
            labels=AGE_BUCKETS
        )
        return df

    in_clean = standardize_df(in_df_raw, "InScan")
    cl_clean = standardize_df(cl_df_raw, "BagClosed")

    all_data = pd.concat([in_clean, cl_clean], ignore_index=True)
    all_data = all_data.drop_duplicates(subset=['bagno'], keep='first')

    def build_summary_table(df, group_col, label_name):
        if df.empty:
            return pd.DataFrame()
        base = df.groupby(group_col).agg(
            InScan_Bags=('Source', lambda x: (x == 'InScan').sum()),
            BagClosed_Bags=('Source', lambda x: (x == 'BagClosed').sum()),
            Total_Bags=('bagno', 'count'),
            Shipments=('totalshipments', 'sum'),
            Weight_Kg=('physicalweight', 'sum')
        )
        age_piv = df.pivot_table(
            index=group_col,
            columns='Age_Bucket',
            values='bagno',
            aggfunc='count',
            fill_value=0,
            observed=False
        )
        for b in AGE_BUCKETS:
            if b not in age_piv.columns:
                age_piv[b] = 0
        age_piv = age_piv[AGE_BUCKETS]
        
        merged = base.join(age_piv).reset_index().rename(columns={
            group_col: label_name,
            'InScan_Bags': 'InScan',
            'BagClosed_Bags': 'BagClosed',
            'Total_Bags': 'Total Bags',
            'Weight_Kg': 'Weight (Kg)'
        })
        merged['Weight (Kg)'] = merged['Weight (Kg)'].round(1)
        merged['Shipments'] = merged['Shipments'].astype(int)
        merged = merged.sort_values(by='Total Bags', ascending=False).reset_index(drop=True)
        
        tot_row = {label_name: '🔥 GRAND TOTAL'}
        for col in ['InScan', 'BagClosed', 'Total Bags', 'Shipments'] + AGE_BUCKETS:
            tot_row[col] = int(merged[col].sum())
        tot_row['Weight (Kg)'] = round(merged['Weight (Kg)'].sum(), 1)
        
        return pd.concat([merged, pd.DataFrame([tot_row])], ignore_index=True)

    conn_all = all_data[all_data['Classification'] == 'Connection']
    proc_all = all_data[all_data['Classification'] == 'Processing']
    unmapped_all = all_data[(all_data['Mapped_Group'] == 'UNMAPPED') | (all_data['Mapped_Lane'] == 'UNMAPPED')]

    st.sidebar.markdown("---")
    with st.sidebar.expander("📲 WhatsApp Live Report", expanded=False):
        c_in = len(conn_all[conn_all['Source'] == 'InScan'])
        c_cl = len(conn_all[conn_all['Source'] == 'BagClosed'])
        p_in = len(proc_all[proc_all['Source'] == 'InScan'])
        p_cl = len(proc_all[proc_all['Source'] == 'BagClosed'])
        wa_text = (
            f"*⚡ PENDENCY PULSE 360 - SRT HUB*\n"
            f"📅 *Time:* {datetime.now().strftime('%d-%b-%Y %I:%M %p')}\n\n"
            f"🚚 *CONNECTION (LINE-HAUL):*\n"
            f"• Total Bags: *{len(conn_all)}* (InScan: {c_in} | Closed: {c_cl})\n"
            f"• Shipments: *{int(conn_all['totalshipments'].sum())}*\n"
            f"• >12 Hrs: *{len(conn_all[conn_all['pendinghrs'] > 12])}* | >24 Hrs: *{len(conn_all[conn_all['pendinghrs'] > 24])}*\n\n"
            f"🏭 *PROCESSING (SURAT LOCAL):*\n"
            f"• Total Bags: *{len(proc_all)}* (InScan: {p_in} | Closed: {p_cl})\n"
            f"• Shipments: *{int(proc_all['totalshipments'].sum())}*\n"
            f"• >12 Hrs: *{len(proc_all[proc_all['pendinghrs'] > 12])}* | >24 Hrs: *{len(proc_all[proc_all['pendinghrs'] > 24])}*\n"
            f"---\n_⚡ Generated via Pendency Pulse (By Jagan)_"
        )
        st.code(wa_text, language="markdown")

    st.sidebar.markdown("---")
    st.sidebar.markdown("<div style='text-align:center; color:#64748b; font-size:12px;'>⚡ <b>Pendency Pulse 360</b><br>Engineered by <b>Jagan</b></div>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs([
        f"🚚 Connection Hub ({len(conn_all)})",
        f"🏭 Processing Hub ({len(proc_all)})",
        f"⚠️ Unmapped Alert ({len(unmapped_all)})",
        "🔍 Instant Bag Radar"
    ])

    DETAIL_COLS = [
        'Source', 'bagno', 'Current_Hub', 'Clean_Dest', 'Mapped_Group',
        'Mapped_Lane', 'pendinghrs', 'Age_Bucket', 'totalshipments',
        'physicalweight', 'Bag_Type', 'Bag_Mode', 'Created_Time',
        'Created_By', 'Last_Modified_By'
    ]

    # --- TAB 1: CONNECTION ---
    with tab1:
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c_in_cnt = len(conn_all[conn_all['Source'] == 'InScan'])
        c_cl_cnt = len(conn_all[conn_all['Source'] == 'BagClosed'])
        c_ship_sum = int(conn_all['totalshipments'].sum())
        c_wt_sum = conn_all['physicalweight'].sum()
        c_12 = len(conn_all[conn_all['pendinghrs'] > 12])
        c_24 = len(conn_all[conn_all['pendinghrs'] > 24])

        c1.markdown(f"<div class='kpi-card kpi-info'><div class='kpi-title'>Total Unique Bags</div><div class='kpi-val'>{len(conn_all)}</div></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='kpi-card kpi-purp'><div class='kpi-title'>InScan / Closed</div><div class='kpi-val'>{c_in_cnt} / {c_cl_cnt}</div></div>", unsafe_allow_html=True)
        c3.markdown(f"<div class='kpi-card kpi-good'><div class='kpi-title'>Total Shipments</div><div class='kpi-val'>{c_ship_sum:,}</div></div>", unsafe_allow_html=True)
        c4.markdown(f"<div class='kpi-card kpi-info'><div class='kpi-title'>Total Weight (Kg)</div><div class='kpi-val'>{c_wt_sum:,.1f}</div></div>", unsafe_allow_html=True)
        c5.markdown(f"<div class='kpi-card kpi-warn'><div class='kpi-title'>Pending > 12 Hrs</div><div class='kpi-val'>{c_12}</div></div>", unsafe_allow_html=True)
        c6.markdown(f"<div class='kpi-card kpi-crit'><div class='kpi-title'>Critical > 24 Hrs</div><div class='kpi-val' style='color:#ef4444;'>{c_24}</div></div>", unsafe_allow_html=True)

        st.markdown("#### 📈 Lane-Wise Connection Matrix (Touch/Click any Row to Open Bag Details)")
        conn_summary = build_summary_table(conn_all, 'Mapped_Lane', 'Lane')
        
        if not conn_summary.empty:
            event_conn = st.dataframe(
                conn_summary,
                use_container_width=True,
                height=380,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )
            st.markdown("<div class='drill-box'><b>🎯 Live Drill-Down Explorer (Connection):</b> Upar table mein kisi bhi Lane par click karein ya niche filter lagayein!</div>", unsafe_allow_html=True)
            
            selected_lane = "ALL LANES"
            if event_conn and len(event_conn.selection.rows) > 0:
                sel_idx = event_conn.selection.rows[0]
                picked = conn_summary.iloc[sel_idx]['Lane']
                if picked != '🔥 GRAND TOTAL':
                    selected_lane = picked

            f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 1.5, 1.2, 1])
            lane_options = ["ALL LANES"] + sorted(conn_all['Mapped_Lane'].unique().tolist())
            l_idx = lane_options.index(selected_lane) if selected_lane in lane_options else 0
            chosen_lane = f_col1.selectbox("1️⃣ Select Lane:", lane_options, index=l_idx, key="c_lane")
            chosen_bucket = f_col2.selectbox("2️⃣ Select Ageing Bucket:", ["ALL BUCKETS"] + AGE_BUCKETS, key="c_bkt")
            chosen_src = f_col3.selectbox("3️⃣ Select Stage:", ["ALL (InScan + Closed)", "InScan", "BagClosed"], key="c_src")

            drill_conn = conn_all.copy()
            if chosen_lane != "ALL LANES":
                drill_conn = drill_conn[drill_conn['Mapped_Lane'] == chosen_lane]
            if chosen_bucket != "ALL BUCKETS":
                drill_conn = drill_conn[drill_conn['Age_Bucket'] == chosen_bucket]
            if chosen_src != "ALL (InScan + Closed)":
                drill_conn = drill_conn[drill_conn['Source'] == chosen_src]

            drill_conn = drill_conn.sort_values(by='pendinghrs', ascending=False)
            with f_col4:
                st.write("")
                st.write("")
                st.download_button("📥 Export Filtered", data=to_excel(drill_conn[DETAIL_COLS]), file_name="Conn_Drilldown.xlsx", use_container_width=True)

            st.caption(f"Showing **{len(drill_conn)}** Bags for **{chosen_lane}** | Bucket: **{chosen_bucket}**")
            st.dataframe(drill_conn[DETAIL_COLS], use_container_width=True, height=260, hide_index=True)

            st.markdown("---")
            st.markdown("#### 🔥 Top 50 High Ageing Actionable Lists (Separate View)")
            t_left, t_right = st.columns(2)
            with t_left:
                st.markdown("##### 🔴 Top 50 Connection - InScan Bags")
                top50_c_in = conn_all[conn_all['Source'] == 'InScan'].sort_values(by='pendinghrs', ascending=False).head(50)
                st.dataframe(top50_c_in[DETAIL_COLS], use_container_width=True, height=350, hide_index=True)
            with t_right:
                st.markdown("##### 🟠 Top 50 Connection - BagClosed Bags")
                top50_c_cl = conn_all[conn_all['Source'] == 'BagClosed'].sort_values(by='pendinghrs', ascending=False).head(50)
                st.dataframe(top50_c_cl[DETAIL_COLS], use_container_width=True, height=350, hide_index=True)

    # --- TAB 2: PROCESSING ---
    with tab2:
        p1, p2, p3, p4, p5, p6 = st.columns(6)
        p_in_cnt = len(proc_all[proc_all['Source'] == 'InScan'])
        p_cl_cnt = len(proc_all[proc_all['Source'] == 'BagClosed'])
        p_ship_sum = int(proc_all['totalshipments'].sum())
        p_wt_sum = proc_all['physicalweight'].sum()
        p_12 = len(proc_all[proc_all['pendinghrs'] > 12])
        p_24 = len(proc_all[proc_all['pendinghrs'] > 24])

        p1.markdown(f"<div class='kpi-card kpi-info'><div class='kpi-title'>Total Processing Bags</div><div class='kpi-val'>{len(proc_all)}</div></div>", unsafe_allow_html=True)
        p2.markdown(f"<div class='kpi-card kpi-purp'><div class='kpi-title'>InScan / Closed</div><div class='kpi-val'>{p_in_cnt} / {p_cl_cnt}</div></div>", unsafe_allow_html=True)
        p3.markdown(f"<div class='kpi-card kpi-good'><div class='kpi-title'>Total Shipments</div><div class='kpi-val'>{p_ship_sum:,}</div></div>", unsafe_allow_html=True)
        p4.markdown(f"<div class='kpi-card kpi-info'><div class='kpi-title'>Total Weight (Kg)</div><div class='kpi-val'>{p_wt_sum:,.1f}</div></div>", unsafe_allow_html=True)
        p5.markdown(f"<div class='kpi-card kpi-warn'><div class='kpi-title'>Pending > 12 Hrs</div><div class='kpi-val'>{p_12}</div></div>", unsafe_allow_html=True)
        p6.markdown(f"<div class='kpi-card kpi-crit'><div class='kpi-title'>Critical > 24 Hrs</div><div class='kpi-val' style='color:#ef4444;'>{p_24}</div></div>", unsafe_allow_html=True)

        b_type_up = proc_all['Bag_Type'].astype(str).str.upper()
        b_mode_up = proc_all['Bag_Mode'].astype(str).str.upper()
        rto_cnt = int(b_type_up.str.contains('RTO', na=False).sum())
        dto_cnt = int(b_type_up.str.contains('DTO', na=False).sum())
        dmg_cnt = int(b_type_up.str.contains('DMG|DAMAGE', na=False).sum())
        air_cnt = int(b_mode_up.str.contains('AIR', na=False).sum())
        surf_cnt = len(proc_all) - air_cnt

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.markdown(f"<div class='mini-card'><div class='mini-title'>🔄 RTO BAGS</div><div class='mini-val'>{rto_cnt}</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='mini-card'><div class='mini-title'>📦 DTO BAGS</div><div class='mini-val'>{dto_cnt}</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='mini-card'><div class='mini-title'>⚠️ DMG (DAMAGE)</div><div class='mini-val'>{dmg_cnt}</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='mini-card'><div class='mini-title'>✈️ AIR MODE</div><div class='mini-val'>{air_cnt}</div></div>", unsafe_allow_html=True)
        m5.markdown(f"<div class='mini-card'><div class='mini-title'>🚛 SURFACE MODE</div><div class='mini-val'>{surf_cnt}</div></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📈 Hub-Wise Processing Matrix (Touch/Click any Row to Open Bag Details)")
        proc_sum
