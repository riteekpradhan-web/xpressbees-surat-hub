import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# --- 1. BIG SCALE PAGE CONFIG ---
st.set_page_config(page_title="Surat Hub Pendency Monitor", page_icon="📦", layout="wide")

st.markdown("""
<style>
    /* Fixed top padding so Header & Tabs never get cut off */
    .block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; max-width: 98% !important; }
    .top-bar {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: white; padding: 18px 26px; border-radius: 12px;
        display: flex; justify-content: space-between; align-items: center;
        margin-bottom: 18px; border-left: 6px solid #3b82f6;
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }
    .top-title { font-size: 26px; font-weight: 800; margin: 0; color: #ffffff; letter-spacing: 0.5px; }
    .top-sub { font-size: 13px; color: #94a3b8; margin-top: 4px; }
    .owner-tag { background: #1e293b; border: 1px solid #3b82f6; color: #60a5fa; padding: 6px 16px; border-radius: 20px; font-size: 13px; font-weight: 700; }
    .kpi-box { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px 10px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.04); margin-bottom: 10px; border-bottom: 4px solid #3b82f6; }
    .kpi-lbl { font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; }
    .kpi-num { font-size: 24px; font-weight: 800; color: #0f172a; margin-top: 4px; }
    .kpi-sub { font-size: 11px; color: #475569; margin-top: 2px; font-weight: 600; }
    .mini-box { background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 10px; padding: 10px; text-align: center; margin-bottom: 12px; }
    .drill-bar { background: #eff6ff; border-left: 5px solid #2563eb; padding: 12px 16px; border-radius: 8px; margin: 12px 0; font-size: 14px; }
    .bag-card { background: #0f172a; color: white; padding: 20px; border-radius: 12px; margin-bottom: 16px; border-left: 6px solid #f59e0b; }
</style>
""", unsafe_allow_html=True)

# --- 2. AUTO-LOGIN GATE ---
if st.query_params.get("auth") != "success":
    st.markdown("<h2 style='text-align:center; margin-top:50px;'>🔒 Surat Hub Pendency Monitor</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        pwd = st.text_input("Enter Hub Password:", type="password")
        if st.button("Login", use_container_width=True):
            if pwd == st.secrets.get("APP_PASSWORD", "SuratHub@2026"):
                st.query_params["auth"] = "success"
                st.rerun()
            else:
                st.error("❌ Galat Password!")
    st.stop()

# --- 3. TOP BANNER ---
st.markdown("""
<div class="top-bar">
    <div>
        <div class="top-title">📦 Surat Hub Pendency Monitor</div>
        <div class="top-sub">Enterprise Data Analytics & Live Operations Intelligence • XpressBees SRT Hub</div>
    </div>
    <div class="owner-tag">⚡ Analytics by Jagan</div>
</div>
""", unsafe_allow_html=True)

# --- 4. SIDEBAR UPLOAD & PERSISTENCE ---
st.sidebar.header("📂 Data Control Center")
map_up = st.sidebar.file_uploader("1. Mappings (Expriment 2.xlsx)", type=['xlsx'])
in_up = st.sidebar.file_uploader("2. InScan CSV", type=['csv'])
cl_up = st.sidebar.file_uploader("3. BagClosed CSV", type=['csv'])

if map_up:
    with open("map.xlsx", "wb") as f: f.write(map_up.getbuffer())
if in_up:
    with open("in.csv", "wb") as f: f.write(in_up.getbuffer())
if cl_up:
    with open("cl.csv", "wb") as f: f.write(cl_up.getbuffer())

def to_excel(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Analysis')
    return out.getvalue()

def get_col(df, names, default="-"):
    cmap = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in cmap:
            return df[cmap[n.lower()]].fillna(default).astype(str)
    return pd.Series([default] * len(df), index=df.index)

# --- 5. DATA ANALYTICS ENGINE ---
if os.path.exists("map.xlsx") and os.path.exists("in.csv") and os.path.exists("cl.csv"):
    @st.cache_data(ttl=60)
    def read_maps(mtime):
        try:
            xls = pd.ExcelFile("map.xlsx")
            d_df = pd.read_excel(xls, "GroupName").fillna("")
            l_df = pd.read_excel(xls, "lANE").fillna("")
            d_map = dict(zip(d_df['ShipmentDestination'].astype(str).str.strip().str.upper(), d_df['GroupName'].astype(str).str.strip()))
            l_map = dict(zip(l_df['Groups'].astype(str).str.strip().str.upper(), l_df['lANE'].astype(str).str.strip()))
            return d_map, l_map
        except Exception:
            return {}, {}

    dest_map, lane_map = read_maps(os.path.getmtime("map.xlsx"))
    BUCKETS = ['0-1 hr', '1-3 hrs', '3-6 hrs', '6-12 hrs', '12-24 hrs', '>24 hrs']

    def clean_df(df, src):
        df = df.copy()
        df['Source'] = src
        df['bagno'] = get_col(df, ['bagno', 'bag_no', 'bagnumber'], 'UNKNOWN').str.strip()
        df['totalshipments'] = pd.to_numeric(get_col(df, ['totalshipments', 'shipments'], 0), errors='coerce').fillna(0)
        df['physicalweight'] = pd.to_numeric(get_col(df, ['physicalweight', 'weight'], 0), errors='coerce').fillna(0)
        df['pendinghrs'] = pd.to_numeric(get_col(df, ['pendinghrs', 'pending_hrs', 'ageing'], 0), errors='coerce').fillna(0)
        
        raw_d = get_col(df, ['destinationhubname', 'destination'], 'UNKNOWN')
        df['Clean_Dest'] = raw_d.apply(lambda x: str(x).split('(')[0].strip().upper())
        df['Current_Hub'] = get_col(df, ['currenthubname', 'currenthub', 'origin'], 'SRT/PC2').str.strip()
        df['Created_Time'] = get_col(df, ['bagcreateddate', 'createddate', 'bagcreatedtime'], '-')
        df['Created_By'] = get_col(df, ['bagcreatedby', 'createdby'], '-')
        df['Last_Modified_By'] = get_col(df, ['lastmodifiedby', 'modifiedby'], '-')
        df['Bag_Type'] = get_col(df, ['bagtype', 'bag_type', 'type'], 'Normal').str.strip()
        df['Bag_Mode'] = get_col(df, ['bagmode', 'mode', 'servicemode'], 'Surface').str.strip()
        
        df['Mapped_Group'] = df['Clean_Dest'].apply(lambda x: dest_map.get(x, "UNMAPPED"))
        df['Mapped_Lane'] = df['Mapped_Group'].apply(lambda x: lane_map.get(x.upper(), "UNMAPPED") if x != "UNMAPPED" else "UNMAPPED")
        
        is_proc = df['Clean_Dest'].str.startswith("SRT") | (df['Mapped_Group'].str.upper() == "SURAT CITY") | (df['Mapped_Lane'].str.upper() == "LOCAL")
        df['Classification'] = is_proc.map({True: "Processing", False: "Connection"})
        df['Age_Bucket'] = pd.cut(df['pendinghrs'], bins=[-1, 1, 3, 6, 12, 24, 999999], labels=BUCKETS)
        return df

    all_data = pd.concat([clean_df(pd.read_csv("in.csv"), "InScan"), clean_df(pd.read_csv("cl.csv"), "BagClosed")], ignore_index=True)
    all_data = all_data.drop_duplicates(subset=['bagno'], keep='first')

    conn_df = all_data[all_data['Classification'] == 'Connection']
    proc_df = all_data[all_data['Classification'] == 'Processing']
    unmap_df = all_data[(all_data['Mapped_Group'] == 'UNMAPPED') | (all_data['Mapped_Lane'] == 'UNMAPPED')]

    with st.sidebar.expander("📲 WhatsApp Executive Summary", expanded=True):
        st.code(
            f"*📦 SURAT HUB PENDENCY*\n🕒 {datetime.now().strftime('%d-%b %I:%M %p')}\n\n"
            f"🚚 *Connection:* {len(conn_df)} Bags ({int(conn_df['totalshipments'].sum()):,} Ships)\n"
            f"   • InScan: {int((conn_df['Source']=='InScan').sum())} | Closed: {int((conn_df['Source']=='BagClosed').sum())}\n"
            f"   • >12h: {int((conn_df['pendinghrs']>12).sum())} | >24h: {int((conn_df['pendinghrs']>24).sum())}\n\n"
            f"🏭 *Processing:* {len(proc_df)} Bags ({int(proc_df['totalshipments'].sum()):,} Ships)\n"
            f"   • InScan: {int((proc_df['Source']=='InScan').sum())} | Closed: {int((proc_df['Source']=='BagClosed').sum())}\n"
            f"   • >12h: {int((proc_df['pendinghrs']>12).sum())} | >24h: {int((proc_df['pendinghrs']>24).sum())}\n"
            f"_Analytics by Jagan_"
        )

    COLS = ['Source', 'bagno', 'Current_Hub', 'Clean_Dest', 'Mapped_Group', 'Mapped_Lane', 'pendinghrs', 'Age_Bucket', 'totalshipments', 'physicalweight', 'Bag_Type', 'Bag_Mode', 'Created_Time', 'Created_By', 'Last_Modified_By']

    def build_matrix(df, grp_col, grp_label):
        base = df.groupby(grp_col).agg(
            InScan=('Source', lambda x: (x == 'InScan').sum()),
            BagClosed=('Source', lambda x: (x == 'BagClosed').sum()),
            Total_Bags=('bagno', 'count'),
            Shipments=('totalshipments', 'sum'),
            Weight_Kg=('physicalweight', 'sum'),
            Avg_Hrs=('pendinghrs', 'mean'),
            Breach_12h=('pendinghrs', lambda x: (x > 12).sum())
        )
        piv = df.pivot_table(index=grp_col, columns='Age_Bucket', values='bagno', aggfunc='count', fill_value=0, observed=False)
        for b in BUCKETS:
            if b not in piv.columns: piv[b] = 0
        m = base.join(piv[BUCKETS]).reset_index().rename(columns={grp_col: grp_label, 'Total_Bags': 'Total Bags', 'Weight_Kg': 'Weight (Kg)', 'Avg_Hrs': 'Avg Hrs'})
        m['>12h %'] = ((m['Breach_12h'] / m['Total Bags']) * 100).round(1).astype(str) + '%'
        m['Weight (Kg)'] = m['Weight (Kg)'].round(1)
        m['Avg Hrs'] = m['Avg Hrs'].round(1)
        m['Shipments'] = m['Shipments'].astype(int)
        m = m.drop(columns=['Breach_12h']).sort_values(by='Total Bags', ascending=False).reset_index(drop=True)

        tot = {grp_label: '🔥 GRAND TOTAL', 'Weight (Kg)': round(df['physicalweight'].sum(), 1), 'Avg Hrs': round(df['pendinghrs'].mean(), 1), '>12h %': f"{round(((df['pendinghrs']>12).sum()/len(df))*100, 1)}%"}
        for c in ['InScan', 'BagClosed', 'Total Bags', 'Shipments'] + BUCKETS:
            tot[c] = int(m[c].sum())
        return pd.concat([m, pd.DataFrame([tot])], ignore_index=True)

    # --- REUSABLE DEEP ANALYST SECTION ---
    def render_section(df, grp_col, grp_label, is_proc, k_pre):
        if df.empty:
            st.warning("No data available.")
            return
        in_c, cl_c = int((df['Source'] == 'InScan').sum()), int((df['Source'] == 'BagClosed').sum())
        ships, wt = int(df['totalshipments'].sum()), df['physicalweight'].sum()
        p12, p24 = int((df['pendinghrs'] > 12).sum()), int((df['pendinghrs'] > 24).sum())
        avg_h = round(df['pendinghrs'].mean(), 1)
        sla_pct = round((p12 / len(df)) * 100, 1)

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Total Unique Bags</div><div class='kpi-num'>{len(df):,}</div><div class='kpi-sub'>Avg {int(ships/len(df))} Ships/Bag</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>InScan vs Closed</div><div class='kpi-num'>{in_c} / {cl_c}</div><div class='kpi-sub'>{round(in_c/len(df)*100)}% InScan Stage</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Total Shipments</div><div class='kpi-num'>{ships:,}</div><div class='kpi-sub'>Active Load</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Physical Weight</div><div class='kpi-num'>{wt:,.1f}</div><div class='kpi-sub'>Avg {round(wt/len(df),1)} Kg/Bag</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='kpi-box' style='border-bottom-color:#f59e0b;'><div class='kpi-lbl'>Pending > 12 Hrs</div><div class='kpi-num' style='color:#d97706;'>{p12:,}</div><div class='kpi-sub'>{sla_pct}% Breach Rate</div></div>", unsafe_allow_html=True)
        k6.markdown(f"<div class='kpi-box' style='border-bottom-color:#ef4444;'><div class='kpi-lbl'>Critical > 24 Hrs</div><div class='kpi-num' style='color:#dc2626;'>{p24:,}</div><div class='kpi-sub'>Avg Age: {avg_h} Hrs</div></div>", unsafe_allow_html=True)

        if is_proc:
            bt, bm = df['Bag_Type'].str.upper(), df['Bag_Mode'].str.upper()
            rto, dto, dmg = int(bt.str.contains('RTO', na=False).sum()), int(bt.str.contains('DTO', na=False).sum()), int(bt.str.contains('DMG|DAMAGE', na=False).sum())
            norm = len(df) - (rto + dto + dmg)
            air = int(bm.str.contains('AIR', na=False).sum())
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.markdown(f"<div class='mini-box'><b>📦 NORMAL:</b> {norm}</div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='mini-box'><b>🔄 RTO BAGS:</b> {rto}</div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='mini-box'><b>📥 DTO BAGS:</b> {dto}</div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='mini-box'><b>⚠️ DAMAGE:</b> {dmg}</div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='mini-box'><b>✈️ AIR MODE:</b> {air}</div>", unsafe_allow_html=True)
            m6.markdown(f"<div class='mini-box'><b>🚛 SURFACE:</b> {len(df)-air}</div>", unsafe_allow_html=True)

        summary_df = build_matrix(df, grp_col, grp_label)
        st.markdown(f"#### 📈 {grp_label}-Wise Master Ageing Matrix (Click any row to inspect bags)")
        ev = st.dataframe(summary_df, use_container_width=True, height=420, hide_index=True, on_select="rerun", selection_mode="single-row", key=f"{k_pre}_tbl")

        # --- DEEP DATA ANALYST BREAKDOWNS ---
        with st.expander("🧠 Deep Data Analyst Breakdown (Origin Hub Load • Bag Type Matrix • Operator Accountability)", expanded=True):
            a1, a2, a3 = st.columns(3)
            with a1:
                st.markdown("##### 🏢 Current Hub / Floor Load")
                h_df = df.groupby('Current_Hub').agg(Bags=('bagno', 'count'), InScan=('Source', lambda x: (x=='InScan').sum()), Closed=('Source', lambda x: (x=='BagClosed').sum()), Ships=('totalshipments', 'sum'), Over24h=('pendinghrs', lambda x: (x>24).sum())).reset_index().sort_values(by='Bags', ascending=False)
                st.dataframe(h_df, use_container_width=True, height=230, hide_index=True)
            with a2:
                st.markdown("##### 📦 Bag Type & Mode Analysis")
                t_df = df.groupby(['Bag_Type', 'Bag_Mode']).agg(Bags=('bagno', 'count'), Ships=('totalshipments', 'sum'), Avg_Hrs=('pendinghrs', lambda x: round(x.mean(),1)), Over24h=('pendinghrs', lambda x: (x>24).sum())).reset_index().sort_values(by='Bags', ascending=False)
                st.dataframe(t_df, use_container_width=True, height=230, hide_index=True)
            with a3:
                st.markdown("##### 👤 Staff / Operator Accountability (Top Pending)")
                u_df = df.groupby('Last_Modified_By').agg(Bags=('bagno', 'count'), Over12h=('pendinghrs', lambda x: (x>12).sum()), Over24h=('pendinghrs', lambda x: (x>24).sum()), Avg_Hrs=('pendinghrs', lambda x: round(x.mean(),1))).reset_index().sort_values(by='Over24h', ascending=False).head(25)
                st.dataframe(u_df, use_container_width=True, height=230, hide_index=True)

        # --- DRILL-DOWN EXPLORER ---
        sel_val = "ALL"
        if ev and len(ev.selection.rows) > 0:
            picked = summary_df.iloc[ev.selection.rows[0]][grp_label]
            if picked != '🔥 GRAND TOTAL': sel_val = picked

        st.markdown("<div class='drill-bar'><b>🎯 Interactive Bag Drill-Down Explorer:</b> Filter by Hub/Lane, Ageing Bucket, Stage, or Bag Type:</div>", unsafe_allow_html=True)
        f1, f2, f3, f4, f5 = st.columns([1.4, 1.2, 1.1, 1.1, 1])
        opts = ["ALL"] + sorted(df[grp_col].unique().tolist())
        c_grp = f1.selectbox(f"1. {grp_label}:", opts, index=opts.index(sel_val) if sel_val in opts else 0, key=f"{k_pre}_g")
        c_bkt = f2.selectbox("2. Ageing Bucket:", ["ALL"] + BUCKETS, key=f"{k_pre}_b")
        c_src = f3.selectbox("3. Stage:", ["ALL", "InScan", "BagClosed"], key=f"{k_pre}_s")
        c_typ = f4.selectbox("4. Bag Type:", ["ALL"] + sorted(df['Bag_Type'].unique().tolist()), key=f"{k_pre}_t")

        sub = df.copy()
        if c_grp != "ALL": sub = sub[sub[grp_col] == c_grp]
        if c_bkt != "ALL": sub = sub[sub['Age_Bucket'] == c_bkt]
        if c_src != "ALL": sub = sub[sub['Source'] == c_src]
        if c_typ != "ALL": sub = sub[sub['Bag_Type'] == c_typ]
        sub = sub.sort_values(by='pendinghrs', ascending=False)

        with f5:
            st.write("")
            st.download_button(f"📥 Export ({len(sub)})", data=to_excel(sub[COLS]), file_name=f"{k_pre}_{c_grp}_bags.xlsx", use_container_width=True, key=f"{k_pre}_dl")

        st.dataframe(sub[COLS], use_container_width=True, height=300, hide_index=True)
        st.markdown("---")
        st.markdown("#### 🔥 Top 50 High Ageing Actionable Bags (InScan vs BagClosed)")
        l_col, r_col = st.columns(2)
        with l_col:
            st.markdown("##### 🔴 Top 50 InScan Bags")
            st.dataframe(df[df['Source'] == 'InScan'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=360, hide_index=True)
        with r_col:
            st.markdown("##### 🟠 Top 50 BagClosed Bags")
            st.dataframe(df[df['Source'] == 'BagClosed'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=360, hide_index=True)

    # --- 6. MAIN TABS ---
    t1, t2, t3, t4 = st.tabs([f"🚚 Connection Hub ({len(conn_df)})", f"🏭 Processing Hub ({len(proc_df)})", f"⚠️ Unmapped ({len(unmap_df)})", "🔍 Deep Bag Search & Kundli"])
    with t1: render_section(conn_df, 'Mapped_Lane', 'Lane', False, "conn")
    with t2: render_section(proc_df, 'Clean_Dest', 'Hub', True, "proc")
    with t3:
        if not unmap_df.empty:
            u_piv = unmap_df.groupby(['Clean_Dest', 'Mapped_Group', 'Mapped_Lane']).agg(Bags=('bagno', 'count'), Shipments=('totalshipments', 'sum'), Weight=('physicalweight', 'sum')).reset_index().sort_values(by='Bags', ascending=False)
            st.warning(f"⚠️ Total {len(u_piv)} unmapped destinations found! Add them to Expriment 2.xlsx.")
            st.dataframe(u_piv, use_container_width=True, hide_index=True)
            st.download_button("📥 Download Unmapped", data=to_excel(u_piv), file_name="Unmapped.xlsx")
        else:
            st.success("✅ All destinations are 100% mapped!")
    with t4:
        st.markdown("#### 🔍 Deep Bag Diagnostic Search (Search by Bag No, Destination, or Staff ID)")
        s_col1, s_col2 = st.columns([3, 1])
        q = s_col1.text_input("📦 Enter Bag Number(s) (comma-separated) or Destination Hub:", placeholder="e.g. nxba17260229 or HYD/VAP")
        s_mode = s_col2.selectbox("Search Mode:", ["By Bag Number", "By Destination Hub", "By Created/Modified User"])
        if q:
            tokens = [t.strip() for t in q.split(",") if t.strip()]
            pat = "|".join(tokens)
            if s_mode == "By Bag Number":
                res = all_data[all_data['bagno'].str.contains(pat, case=False, na=False)]
            elif s_mode == "By Destination Hub":
                res = all_data[all_data['Clean_Dest'].str.contains(pat, case=False, na=False)]
            else:
                res = all_data[all_data['Created_By'].str.contains(pat, case=False, na=False) | all_data['Last_Modified_By'].str.contains(pat, case=False, na=False)]

            if not res.empty:
                st.success(f"✅ Found **{len(res)}** matching bag(s)!")
                # Deep 360 Kundli Card for the top/searched bag
                b = res.iloc[0]
                days_p = round(float(b['pendinghrs']) / 24, 1)
                sla_tag = "🚨 CRITICAL SLA BREACH (>24h)" if float(b['pendinghrs']) > 24 else ("⚠️ WARNING (>12h)" if float(b['pendinghrs']) > 12 else "✅ WITHIN SLA")
                st.markdown(f"""
                <div class='bag-card'>
                    <div style='display:flex; justify-content:space-between; align-items:center;'>
                        <h3 style='margin:0; color:#fbbf24;'>📦 Bag Kundli: {b['bagno']} ({b['Source']})</h3>
                        <span style='background:#ef4444; color:white; padding:4px 12px; borde
