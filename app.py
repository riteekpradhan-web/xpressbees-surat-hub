import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# --- 1. PAGE CONFIG ---
st.set_page_config(page_title="Surat Hub Pendency Monitor", page_icon="📦", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 98%; }
    .top-bar {
        background: #0f172a; color: white; padding: 14px 22px; border-radius: 10px;
        display: flex; justify-content: space-between; align-items: center;
        margin-bottom: 16px; border-left: 5px solid #3b82f6;
    }
    .top-title { font-size: 22px; font-weight: 700; margin: 0; color: #fff; }
    .top-sub { font-size: 12px; color: #94a3b8; }
    .owner-tag { background: #1e293b; border: 1px solid #475569; color: #e2e8f0; padding: 4px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; }
    .kpi-box { background: #fff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.04); margin-bottom: 8px; }
    .kpi-lbl { font-size: 11px; font-weight: 600; color: #64748b; text-transform: uppercase; }
    .kpi-num { font-size: 22px; font-weight: 800; color: #0f172a; margin-top: 2px; }
    .mini-box { background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px; text-align: center; margin-bottom: 10px; }
    .drill-bar { background: #eff6ff; border-left: 4px solid #2563eb; padding: 10px 14px; border-radius: 6px; margin: 8px 0; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# --- 2. AUTO-LOGIN GATE ---
if st.query_params.get("auth") != "success":
    st.markdown("<h3 style='text-align:center; margin-top:60px;'>🔒 Surat Hub Pendency Monitor</h3>", unsafe_allow_html=True)
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

# --- 3. HEADER ---
st.markdown("""
<div class="top-bar">
    <div>
        <div class="top-title">📦 Surat Hub Pendency Monitor</div>
        <div class="top-sub">XpressBees Live Connection & Processing Operations</div>
    </div>
    <div class="owner-tag">👤 Jagan</div>
</div>
""", unsafe_allow_html=True)

# --- 4. SIDEBAR UPLOAD & SAVE ---
st.sidebar.header("📂 Data Upload")
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
        df.to_excel(writer, index=False, sheet_name='Data')
    return out.getvalue()

def get_col(df, names, default="-"):
    cmap = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in cmap:
            return df[cmap[n.lower()]].fillna(default).astype(str)
    return pd.Series([default] * len(df), index=df.index)

# --- 5. DATA PROCESSING ---
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
        df['Current_Hub'] = get_col(df, ['currenthubname', 'currenthub', 'origin'], 'SRT/PC2')
        df['Created_Time'] = get_col(df, ['bagcreateddate', 'createddate', 'bagcreatedtime'], '-')
        df['Created_By'] = get_col(df, ['bagcreatedby', 'createdby'], '-')
        df['Last_Modified_By'] = get_col(df, ['lastmodifiedby', 'modifiedby'], '-')
        df['Bag_Type'] = get_col(df, ['bagtype', 'bag_type', 'type'], 'Normal')
        df['Bag_Mode'] = get_col(df, ['bagmode', 'mode', 'servicemode'], 'Surface')
        
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

    # WhatsApp Summary in Sidebar
    with st.sidebar.expander("📲 WhatsApp Summary"):
        st.code(
            f"*📦 SURAT HUB PENDENCY*\n🕒 {datetime.now().strftime('%d-%b %I:%M %p')}\n\n"
            f"🚚 *Connection:* {len(conn_df)} Bags ({int(conn_df['totalshipments'].sum())} Ships) | >24h: {len(conn_df[conn_df['pendinghrs']>24])}\n"
            f"🏭 *Processing:* {len(proc_df)} Bags ({int(proc_df['totalshipments'].sum())} Ships) | >24h: {len(proc_df[proc_df['pendinghrs']>24])}\n"
            f"_By Jagan_"
        )

    COLS = ['Source', 'bagno', 'Current_Hub', 'Clean_Dest', 'Mapped_Group', 'Mapped_Lane', 'pendinghrs', 'Age_Bucket', 'totalshipments', 'physicalweight', 'Bag_Type', 'Bag_Mode', 'Created_Time', 'Created_By', 'Last_Modified_By']

    # --- REUSABLE DASHBOARD RENDERER ---
    def render_section(df, grp_col, grp_label, is_proc, k_pre):
        if df.empty:
            st.warning("No data available.")
            return
        in_c = int((df['Source'] == 'InScan').sum())
        cl_c = int((df['Source'] == 'BagClosed').sum())
        ships = int(df['totalshipments'].sum())
        wt = df['physicalweight'].sum()
        p12 = int((df['pendinghrs'] > 12).sum())
        p24 = int((df['pendinghrs'] > 24).sum())

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Total Bags</div><div class='kpi-num'>{len(df)}</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>InScan / Closed</div><div class='kpi-num'>{in_c} / {cl_c}</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Shipments</div><div class='kpi-num'>{ships:,}</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Weight (Kg)</div><div class='kpi-num'>{wt:,.1f}</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Pending > 12h</div><div class='kpi-num' style='color:#d97706;'>{p12}</div></div>", unsafe_allow_html=True)
        k6.markdown(f"<div class='kpi-box'><div class='kpi-lbl'>Critical > 24h</div><div class='kpi-num' style='color:#dc2626;'>{p24}</div></div>", unsafe_allow_html=True)

        if is_proc:
            bt = df['Bag_Type'].str.upper()
            bm = df['Bag_Mode'].str.upper()
            rto, dto, dmg = int(bt.str.contains('RTO', na=False).sum()), int(bt.str.contains('DTO', na=False).sum()), int(bt.str.contains('DMG|DAMAGE', na=False).sum())
            air = int(bm.str.contains('AIR', na=False).sum())
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.markdown(f"<div class='mini-box'><b>🔄 RTO:</b> {rto}</div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='mini-box'><b>📦 DTO:</b> {dto}</div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='mini-box'><b>⚠️ DMG:</b> {dmg}</div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='mini-box'><b>✈️ AIR:</b> {air}</div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='mini-box'><b>🚛 SURFACE:</b> {len(df)-air}</div>", unsafe_allow_html=True)

        # Build Summary Table
        base = df.groupby(grp_col).agg(
            InScan=('Source', lambda x: (x == 'InScan').sum()),
            BagClosed=('Source', lambda x: (x == 'BagClosed').sum()),
            Total_Bags=('bagno', 'count'),
            Shipments=('totalshipments', 'sum'),
            Weight_Kg=('physicalweight', 'sum')
        )
        piv = df.pivot_table(index=grp_col, columns='Age_Bucket', values='bagno', aggfunc='count', fill_value=0, observed=False)
        for b in BUCKETS:
            if b not in piv.columns: piv[b] = 0
        merged = base.join(piv[BUCKETS]).reset_index().rename(columns={grp_col: grp_label, 'Total_Bags': 'Total Bags', 'Weight_Kg': 'Weight (Kg)'})
        merged['Weight (Kg)'] = merged['Weight (Kg)'].round(1)
        merged['Shipments'] = merged['Shipments'].astype(int)
        merged = merged.sort_values(by='Total Bags', ascending=False).reset_index(drop=True)

        tot = {grp_label: '🔥 GRAND TOTAL', 'Weight (Kg)': round(merged['Weight (Kg)'].sum(), 1)}
        for c in ['InScan', 'BagClosed', 'Total Bags', 'Shipments'] + BUCKETS:
            tot[c] = int(merged[c].sum())
        summary_df = pd.concat([merged, pd.DataFrame([tot])], ignore_index=True)

        st.markdown(f"#### 📈 {grp_label}-Wise Pendency Summary (Click any row to view bags below)")
        ev = st.dataframe(summary_df, use_container_width=True, height=350, hide_index=True, on_select="rerun", selection_mode="single-row", key=f"{k_pre}_tbl")

        # Drill-down Explorer
        sel_val = "ALL"
        if ev and len(ev.selection.rows) > 0:
            picked = summary_df.iloc[ev.selection.rows[0]][grp_label]
            if picked != '🔥 GRAND TOTAL': sel_val = picked

        st.markdown("<div class='drill-bar'><b>🎯 Bag Drill-Down Explorer:</b> Upar table ki row par click karein ya niche filter chunein:</div>", unsafe_allow_html=True)
        f1, f2, f3, f4 = st.columns([1.5, 1.5, 1.2, 1])
        opts = ["ALL"] + sorted(df[grp_col].unique().tolist())
        c_grp = f1.selectbox(f"Select {grp_label}:", opts, index=opts.index(sel_val) if sel_val in opts else 0, key=f"{k_pre}_g")
        c_bkt = f2.selectbox("Select Ageing Bucket:", ["ALL"] + BUCKETS, key=f"{k_pre}_b")
        c_src = f3.selectbox("Select Stage:", ["ALL", "InScan", "BagClosed"], key=f"{k_pre}_s")

        sub = df.copy()
        if c_grp != "ALL": sub = sub[sub[grp_col] == c_grp]
        if c_bkt != "ALL": sub = sub[sub['Age_Bucket'] == c_bkt]
        if c_src != "ALL": sub = sub[sub['Source'] == c_src]
        sub = sub.sort_values(by='pendinghrs', ascending=False)

        with f4:
            st.write("")
            st.download_button("📥 Download Excel", data=to_excel(sub[COLS]), file_name=f"{k_pre}_bags.xlsx", use_container_width=True, key=f"{k_pre}_dl")

        st.dataframe(sub[COLS], use_container_width=True, height=250, hide_index=True)
        st.markdown("---")
        st.markdown("#### 🔥 Top 50 High Ageing Bags (InScan vs BagClosed)")
        l_col, r_col = st.columns(2)
        with l_col:
            st.markdown("##### 🔴 Top 50 InScan Bags")
            st.dataframe(df[df['Source'] == 'InScan'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=320, hide_index=True)
        with r_col:
            st.markdown("##### 🟠 Top 50 BagClosed Bags")
            st.dataframe(df[df['Source'] == 'BagClosed'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=320, hide_index=True)

    # --- 6. TABS ---
    t1, t2, t3, t4 = st.tabs([f"🚚 Connection ({len(conn_df)})", f"🏭 Processing ({len(proc_df)})", f"⚠️ Unmapped ({len(unmap_df)})", "🔍 Bag Search"])
    with t1: render_section(conn_df, 'Mapped_Lane', 'Lane', False, "conn")
    with t2: render_section(proc_df, 'Clean_Dest', 'Hub', True, "proc")
    with t3:
        if not unmap_df.empty:
            u_piv = unmap_df.groupby(['Clean_Dest', 'Mapped_Group', 'Mapped_Lane']).agg(Bags=('bagno', 'count'), Shipments=('totalshipments', 'sum')).reset_index().sort_values(by='Bags', ascending=False)
            st.warning(f"Total {len(u_piv)} unmapped destinations found!")
            st.dataframe(u_piv, use_container_width=True, hide_index=True)
        else:
            st.success("✅ All destinations are 100% mapped!")
    with t4:
        q = st.text_input("Enter Bag Number:")
        if q:
            res = all_data[all_data['bagno'].str.contains(q.strip(), case=False, na=False)]
            st.dataframe(res[COLS], use_container_width=True, hide_index=True) if not res.empty else st.error("Bag not found!")
else:
    st.info("👈 Left sidebar se Mapping, InScan, aur BagClosed files ek baar upload karein.")
# --- END OF FILE ---
