import streamlit as st
import pandas as pd
import io, os
from datetime import datetime

st.set_page_config(page_title="Surat Hub Pendency Monitor", page_icon="📦", layout="wide")

# --- 1. SINGLE-LINE CSS (NO TRIPLE QUOTES, HEADER CUT-OFF FIXED) ---
css = "<style>.block-container{padding-top:3.5rem!important;padding-bottom:2rem!important;max-width:98%!important;}.top-bar{background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;padding:16px 24px;border-radius:12px;display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;border-left:6px solid #3b82f6;box-shadow:0 4px 12px rgba(0,0,0,0.12);}.top-title{font-size:24px;font-weight:800;margin:0;color:#fff;}.top-sub{font-size:13px;color:#94a3b8;margin-top:2px;}.owner-tag{background:#1e293b;border:1px solid #3b82f6;color:#60a5fa;padding:6px 14px;border-radius:20px;font-size:13px;font-weight:700;}.kpi-box{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px 8px;text-align:center;box-shadow:0 2px 5px rgba(0,0,0,0.04);margin-bottom:10px;border-bottom:4px solid #3b82f6;}.kpi-lbl{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;}.kpi-num{font-size:23px;font-weight:800;color:#0f172a;margin-top:2px;}.kpi-sub{font-size:11px;color:#475569;font-weight:600;}.mini-box{background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:8px;text-align:center;margin-bottom:10px;}.drill-bar{background:#eff6ff;border-left:5px solid #2563eb;padding:10px 14px;border-radius:6px;margin:10px 0;font-size:14px;}.bag-card{background:#0f172a;color:#fff;padding:16px 20px;border-radius:12px;margin-bottom:14px;border-left:6px solid #f59e0b;}</style>"
st.markdown(css, unsafe_allow_html=True)

# --- 2. AUTO-LOGIN GATE ---
if st.query_params.get("auth") != "success":
    st.markdown("<h3 style='text-align:center;margin-top:40px;'>🔒 Surat Hub Pendency Monitor</h3>", unsafe_allow_html=True)
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
st.markdown("<div class='top-bar'><div><div class='top-title'>📦 Surat Hub Pendency Monitor</div><div class='top-sub'>Enterprise Data Analytics & Live Operations Intelligence • XpressBees SRT Hub</div></div><div class='owner-tag'>⚡ Analytics by Jagan</div></div>", unsafe_allow_html=True)

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
    unmap_df = all_data[(all_data['Mapped_Group'] == 'UNMAPPED') | (all_data['Mapped_Lane'] == 'UNMAPPED
