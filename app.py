import streamlit as st
import pandas as pd
import io, os
from datetime import datetime

st.set_page_config(page_title="Surat Hub Pendency Monitor", page_icon="📦", layout="wide")

# --- ENTERPRISE SCREENSHOT-STYLE CSS ---
css = """
<style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; max-width: 98% !important; }
    .top-bar {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: white; padding: 18px 24px; border-radius: 12px;
        display: flex; justify-content: space-between; align-items: center;
        margin-bottom: 20px; border-left: 6px solid #2563eb;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .top-title { font-size: 24px; font-weight: 800; margin: 0; color: #fff; }
    .top-sub { font-size: 13px; color: #94a3b8; margin-top: 3px; }
    .owner-tag {
        background: #1e293b; border: 1px solid #3b82f6; color: #60a5fa;
        padding: 6px 16px; border-radius: 20px; font-size: 13px; font-weight: 700;
    }
    /* Enterprise KPI Card Style (Matching Screenshot) */
    .kpi-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px;
        padding: 16px 12px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        margin-bottom: 12px; border-top: 3px solid #2563eb;
    }
    .kpi-title { font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-val { font-size: 26px; font-weight: 800; color: #0f172a; margin-top: 6px; }
    .kpi-footer { font-size: 11px; color: #475569; margin-top: 4px; font-weight: 600; }
    .mini-card {
        background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px;
        padding: 10px; text-align: center; margin-bottom: 10px;
    }
    .drill-box {
        background: #eff6ff; border-left: 5px solid #2563eb;
        padding: 12px 16px; border-radius: 8px; margin: 12px 0; font-size: 14px;
    }
    .bag-card {
        background: #0f172a; color: #fff; padding: 18px 22px;
        border-radius: 10px; margin-bottom: 14px; border-left: 6px solid #f59e0b;
    }
</style>
"""
st.markdown(css, unsafe_allow_html=True)

# --- LOGIN GATE ---
if st.query_params.get("auth") != "success":
    st.markdown("<h3 style='text-align:center; margin-top:50px;'>🔒 Surat Hub Pendency Monitor</h3>", unsafe_allow_html=True)
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

# --- HEADER ---
st.markdown("""
<div class="top-bar">
    <div>
        <div class="top-title">📦 Surat Hub Pendency Monitor</div>
        <div class="top-sub">Enterprise Analytics & Live Operations Command Center • XpressBees SRT Hub</div>
    </div>
    <div class="owner-tag">⚡ Analytics by Jagan</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR CONTROL CENTER ---
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
        df.to_excel(writer, index=False, sheet_name='Report')
    return out.getvalue()

def get_col(df, names, default="-"):
    cmap = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in cmap:
            return df[cmap[n.lower()]].fillna(default).astype(str)
    return pd.Series([default] * len(df), index=df.index)

# --- DATA ENGINE ---
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

    with st.sidebar.expander("📲 WhatsApp Summary", expanded=True):
        st.code(f"*📦 SURAT HUB PENDENCY*\n🕒 {datetime.now().strftime('%d-%b %I:%M %p')}\n\n🚚 *Connection:* {len(conn_df)} Bags ({int(conn_df['totalshipments'].sum()):,} Ships)\n🏭 *Processing:* {len(proc_df)} Bags ({int(proc_df['totalshipments'].sum()):,} Ships)\n_Analytics by Jagan_")

    COLS = ['Source', 'bagno', 'Current_Hub', 'Clean_Dest', 'Mapped_Group', 'Mapped_Lane', 'pendinghrs', 'Age_Bucket', 'totalshipments', 'physicalweight', 'Bag_Type', 'Bag_Mode', 'Created_Time', 'Created_By', 'Last_Modified_By']

    def render_section(df, grp_col, grp_label, is_proc, k_pre):
        if df.empty:
            st.warning("No data available.")
            return
        in_c, cl_c = int((df['Source'] == 'InScan').sum()), int((df['Source'] == 'BagClosed').sum())
        ships, wt = int(df['totalshipments'].sum()), df['physicalweight'].sum()
        p12, p24 = int((df['pendinghrs'] > 12).sum()), int((df['pendinghrs'] > 24).sum())
        avg_h, sla_pct = round(df['pendinghrs'].mean(), 1), round((p12 / len(df)) * 100, 1)

        # Screenshot Style KPI Grid Cards
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f"<div class='kpi-card' style='border-top-color:#3b82f6;'><div class='kpi-title'>Total Bags</div><div class='kpi-val'>{len(df):,}</div><div class='kpi-footer'>Avg {int(ships/len(df))} Ships/Bag</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='kpi-card' style='border-top-color:#8b5cf6;'><div class='kpi-title'>InScan / Closed</div><div class='kpi-val'>{in_c} / {cl_c}</div><div class='kpi-footer'>{round(in_c/len(df)*100)}% InScan</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='kpi-card' style='border-top-color:#10b981;'><div class='kpi-title'>Shipments</div><div class='kpi-val'>{ships:,}</div><div class='kpi-footer'>Active Load</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='kpi-card' style='border-top-color:#0ea5e9;'><div class='kpi-title'>Weight (Kg)</div><div class='kpi-val'>{wt:,.1f}</div><div class='kpi-footer'>Avg {round(wt/len(df),1)} Kg</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='kpi-card' style='border-top-color:#f59e0b;'><div class='kpi-title'>Pending > 12h</div><div class='kpi-val' style='color:#d97706;'>{p12:,}</div><div class='kpi-footer'>{sla_pct}% Breach</div></div>", unsafe_allow_html=True)
        k6.markdown(f"<div class='kpi-card' style='border-top-color:#ef4444;'><div class='kpi-title'>Critical > 24h</div><div class='kpi-val' style='color:#dc2626;'>{p24:,}</div><div class='kpi-footer'>Avg {avg_h} Hrs</div></div>", unsafe_allow_html=True)

        if is_proc:
            bt, bm = df['Bag_Type'].str.upper(), df['Bag_Mode'].str.upper()
            rto, dto, dmg = int(bt.str.contains('RTO', na=False).sum()), int(bt.str.contains('DTO', na=False).sum()), int(bt.str.contains('DMG|DAMAGE', na=False).sum())
            air = int(bm.str.contains('AIR', na=False).sum())
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.markdown(f"<div class='mini-box'><b>📦 NORMAL:</b> {len(df)-(rto+dto+dmg)}</div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='mini-box'><b>🔄 RTO:</b> {rto}</div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='mini-box'><b>📥 DTO:</b> {dto}</div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='mini-box'><b>⚠️ DAMAGE:</b> {dmg}</div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='mini-box'><b>✈️ AIR:</b> {air}</div>", unsafe_allow_html=True)
            m6.markdown(f"<div class='mini-box'><b>🚛 SURFACE:</b> {len(df)-air}</div>", unsafe_allow_html=True)

        base = df.groupby(grp_col).agg(InScan=('Source', lambda x: (x=='InScan').sum()), BagClosed=('Source', lambda x: (x=='BagClosed').sum()), Total_Bags=('bagno', 'count'), Shipments=('totalshipments', 'sum'), Weight_Kg=('physicalweight', 'sum'), Avg_Hrs=('pendinghrs', 'mean'), B12=('pendinghrs', lambda x: (x>12).sum()))
        piv = df.pivot_table(index=grp_col, columns='Age_Bucket', values='bagno', aggfunc='count', fill_value=0, observed=False)
        for b in BUCKETS:
            if b not in piv.columns: piv[b] = 0
        m = base.join(piv[BUCKETS]).reset_index().rename(columns={grp_col: grp_label, 'Total_Bags': 'Total Bags', 'Weight_Kg': 'Weight (Kg)', 'Avg_Hrs': 'Avg Hrs'})
        m['>12h %'] = ((m['B12'] / m['Total Bags']) * 100).round(1).astype(str) + '%'
        m['Weight (Kg)'], m['Avg Hrs'], m['Shipments'] = m['Weight (Kg)'].round(1), m['Avg Hrs'].round(1), m['Shipments'].astype(int)
        m = m.drop(columns=['B12']).sort_values(by='Total Bags', ascending=False).reset_index(drop=True)
        tot = {grp_label: '🔥 GRAND TOTAL', 'Weight (Kg)': round(wt, 1), 'Avg Hrs': avg_h, '>12h %': f"{sla_pct}%"}
        for c in ['InScan', 'BagClosed', 'Total Bags', 'Shipments'] + BUCKETS: tot[c] = int(m[c].sum())
        summary_df = pd.concat([m, pd.DataFrame([tot])], ignore_index=True)

        st.markdown(f"#### 📈 {grp_label}-Wise Master Ageing Matrix (Click any row to inspect bags)")
        ev = st.dataframe(summary_df, use_container_width=True, height=380, hide_index=True, on_select="rerun", selection_mode="single-row", key=f"{k_pre}_tbl")

        with st.expander("🧠 Deep Data Analyst Breakdown (Floor Load • Bag Type Matrix • Operator Accountability)", expanded=True):
            a1, a2, a3 = st.columns(3)
            with a1:
                st.markdown("##### 🏢 Current Hub / Floor Load")
                st.dataframe(df.groupby('Current_Hub').agg(Bags=('bagno', 'count'), InScan=('Source', lambda x: (x=='InScan').sum()), Closed=('Source', lambda x: (x=='BagClosed').sum()), Ships=('totalshipments', 'sum'), Over24h=('pendinghrs', lambda x: (x>24).sum())).reset_index().sort_values(by='Bags', ascending=False), use_container_width=True, height=220, hide_index=True)
            with a2:
                st.markdown("##### 📦 Bag Type & Mode Matrix")
                st.dataframe(df.groupby(['Bag_Type', 'Bag_Mode']).agg(Bags=('bagno', 'count'), Ships=('totalshipments', 'sum'), Avg_Hrs=('pendinghrs', lambda x: round(x.mean(),1)), Over24h=('pendinghrs', lambda x: (x>24).sum())).reset_index().sort_values(by='Bags', ascending=False), use_container_width=True, height=220, hide_index=True)
            with a3:
                st.markdown("##### 👤 Operator Accountability (Top >24h)")
                st.dataframe(df.groupby('Last_Modified_By').agg(Bags=('bagno', 'count'), Over12h=('pendinghrs', lambda x: (x>12).sum()), Over24h=('pendinghrs', lambda x: (x>24).sum()), Avg_Hrs=('pendinghrs', lambda x: round(x.mean(),1))).reset_index().sort_values(by='Over24h', ascending=False).head(25), use_container_width=True, height=220, hide_index=True)

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
            st.download_button(f"📥 Export ({len(sub)})", data=to_excel(sub[COLS]), file_name=f"{k_pre}_{c_grp}.xlsx", use_container_width=True, key=f"{k_pre}_dl")

        st.dataframe(sub[COLS], use_container_width=True, height=280, hide_index=True)
        st.markdown("---")
        st.markdown("#### 🔥 Top 50 High Ageing Actionable Bags (InScan vs BagClosed)")
        l_col, r_col = st.columns(2)
        with l_col:
            st.markdown("##### 🔴 Top 50 InScan Bags")
            st.dataframe(df[df['Source'] == 'InScan'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=340, hide_index=True)
        with r_col:
            st.markdown("##### 🟠 Top 50 BagClosed Bags")
            st.dataframe(df[df['Source'] == 'BagClosed'].sort_values(by='pendinghrs', ascending=False).head(50)[COLS], use_container_width=True, height=340, hide_index=True)

    # --- TABS ---
    t1, t2, t3, t4 = st.tabs([f"🚚 Connection Hub ({len(conn_df)})", f"🏭 Processing Hub ({len(proc_df)})", f"⚠️ Unmapped ({len(unmap_df)})", "🔍 Deep Bag Search & Kundli"])
    with t1: render_section(conn_df, 'Mapped_Lane', 'Lane', False, "conn")
    with t2: render_section(proc_df, 'Clean_Dest', 'Hub', True, "proc")
    with t3:
        if not unmap_df.empty:
            u_piv = unmap_df.groupby(['Clean_Dest', 'Mapped_Group', 'Mapped_Lane']).agg(Bags=('bagno', 'count'), Shipments=('totalshipments', 'sum'), Weight=('physicalweight', 'sum')).reset_index().sort_values(by='Bags', ascending=False)
            st.warning(f"⚠️ Total {len(u_piv)} unmapped destinations found!")
            st.dataframe(u_piv, use_container_width=True, hide_index=True)
            st.download_button("📥 Download Unmapped", data=to_excel(u_piv), file_name="Unmapped.xlsx")
        else:
            st.success("✅ All destinations are 100% mapped!")
    with t4:
        st.markdown("#### 🔍 Deep Bag Diagnostic Search (Search by Bag No, Destination, or Staff ID)")
        s1, s2 = st.columns([3, 1])
        q = s1.text_input("📦 Enter Bag Number(s) or Destination Hub or Staff ID:", placeholder="e.g. nxba17260229 or HYD/VAP")
        s_mode = s2.selectbox("Search By:", ["Bag Number", "Destination Hub", "Staff ID"])
        if q:
            pat = "|".join([t.strip() for t in q.split(",") if t.strip()])
            if s_mode == "Bag Number": res = all_data[all_data['bagno'].str.contains(pat, case=False, na=False)]
            elif s_mode == "Destination Hub": res = all_data[all_data['Clean_Dest'].str.contains(pat, case=False, na=False)]
            else: res = all_data[all_data['Created_By'].str.contains(pat, case=False, na=False) | all_data['Last_Modified_By'].str.contains(pat, case=False, na=False)]

            if not res.empty:
                b = res.iloc[0]
                days_p = round(float(b['pendinghrs']) / 24, 1)
                sla = "🚨 CRITICAL SLA BREACH (>24h)" if float(b['pendinghrs']) > 24 else ("⚠️ WARNING (>12h)" if float(b['pendinghrs']) > 12 else "✅ WITHIN SLA")
                card = f"<div class='bag-card'><h4 style='margin:0;color:#fbbf24;'>📦 Bag Kundli: {b['bagno']} ({b['Source']}) — {sla}</h4><hr style='border-color:#334155;margin:8px 0;'><p style='margin:4px 0;font-size:14px;'><b>🛣️ Full Route:</b> Current Hub: <code>{b['Current_Hub']}</code> ➔ Dest: <code>{b['Clean_Dest']}</code> ➔ Group: <code>{b['Mapped_Group']}</code> ➔ Lane: <code>{b['Mapped_Lane']}</code> (<b>{b['Classification']}</b>)</p><p style='margin:4px 0;font-size:14px;'><b>⏱️ Ageing & Load:</b> Pending: <b>{b['pendinghrs']} Hrs ({days_p} Days)</b> | Bucket: <b>{b['Age_Bucket']}</b> | Shipments: <b>{int(b['totalshipments'])}</b> | Weight: <b>{b['physicalweight']} Kg</b> | Type: <b>{b['Bag_Type']} ({b['Bag_Mode']})</b></p><p style='margin:4px 0;font-size:14px;'><b>👤 Audit Trail:</b> Created Time: <code>{b['Created_Time']}</code> | Created By: <code>{b['Created_By']}</code> | Last Modified By: <code>{b['Last_Modified_Body']}</code></p></div>"
                st.markdown(card, unsafe_allow_html=True)
                st.dataframe(res[COLS], use_container_width=True, hide_index=True)
                td = b['Clean_Dest']
                peers = all_data[all_data['Clean_Dest'] == td].sort_values(by='pendinghrs', ascending=False)
                st.markdown(f"##### 📊 Lot Analysis: All Pending Bags for `{td}` ({len(peers)} Bags | {int(peers['totalshipments'].sum())} Shipments)")
                st.dataframe(peers[COLS], use_container_width=True, height=250, hide_index=True)
            else:
                st.error("❌ Record not found!")
else:
    st.info("👈 Left sidebar se Mapping, InScan, aur BagClosed files ek baar upload karein.")
#
