import streamlit as st
import streamlit as st
import pandas as pd

# --- SECURITY GATE (LOGIN SCREEN) ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 XpressBees Secure Portal")
    pwd = st.text_input("Enter Hub Password:", type="password")
    if st.button("Login"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ Galat Password! Access Denied.")
    st.stop()  # Ye line aage ka dashboard tab tak rok degi jab tak login na ho
# ------------------------------------

# (Tera baaki ka purana code iske niche rahega...)
st.set_page_config(page_title="XpressBees Surat Dashboard", layout="wide")
# ...
import pandas as pd
import numpy as np

# --- PAGE SETUP ---
st.set_page_config(page_title="XpressBees Surat Dashboard", layout="wide")
st.title("👑 XpressBees Surat Hub - Deep Analysis Dashboard")
st.markdown("**Owner: Pandi Jagan** | Live Operations Portal")

# --- SIDEBAR UPLOADS ---
st.sidebar.header("📂 Data Upload (Daily)")
mapping_file = st.sidebar.file_uploader("1. Upload Mappings (Expriment 2.xlsx)", type=['xlsx'])
inscan_file = st.sidebar.file_uploader("2. Upload InScan CSV", type=['csv'])
closed_file = st.sidebar.file_uploader("3. Upload BagClosed CSV", type=['csv'])

# --- MAPPING CACHE FUNCTION ---
@st.cache_data
def load_mappings(file):
    try:
        xls = pd.ExcelFile(file)
        dest_df = pd.read_excel(xls, "GroupName").fillna("")
        lane_df = pd.read_excel(xls, "lANE").fillna("")
        dest_map = dict(zip(dest_df['ShipmentDestination'].astype(str).str.strip().str.upper(), dest_df['GroupName'].astype(str).str.strip()))
        lane_map = dict(zip(lane_df['Groups'].astype(str).str.strip().str.upper(), lane_df['lANE'].astype(str).str.strip()))
        return dest_map, lane_map
    except Exception as e:
        st.sidebar.error("Mapping file error. Please check Expriment 2.xlsx")
        return {}, {}

# --- PROCESSING FUNCTION ---
def process_data(df, dest_map, lane_map):
    df = df.copy()
    # Ensure numerical columns are numeric
    for col in ['totalshipments', 'physicalweight', 'pendinghrs']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Clean Destination (Remove brackets)
    if 'destinationhubname' in df.columns:
        df['Clean_Dest'] = df['destinationhubname'].astype(str).apply(lambda x: x.split('(')[0].strip())
    else:
        df['Clean_Dest'] = "Unknown"
        
    # Apply Mappings
    df['Mapped_Group'] = df['Clean_Dest'].apply(lambda x: dest_map.get(x.upper(), "Unmapped"))
    df['Mapped_Lane'] = df['Mapped_Group'].apply(lambda x: lane_map.get(x.upper(), "Unmapped") if x != "Unmapped" else "Unmapped")
    
    # Classification Logic (Processing vs Connection)
    def classify(row):
        dest = str(row.get('Clean_Dest', '')).upper()
        grp = str(row.get('Mapped_Group', '')).upper()
        lane = str(row.get('Mapped_Lane', '')).upper()
        
        if dest.startswith("SRT") or grp == "SURAT CITY" or lane == "LOCAL":
            return "Processing"
        return "Connection"
    
    df['Classification'] = df.apply(classify, axis=1)
    return df

# --- MAIN APP LOGIC ---
if mapping_file and inscan_file and closed_file:
    dest_map, lane_map = load_mappings(mapping_file)
    
    with st.spinner("Analyzing Bags... Please wait!"):
        inscan_df = pd.read_csv(inscan_file)
        closed_df = pd.read_csv(closed_file)
        
        in_proc = process_data(inscan_df, dest_map, lane_map)
        cl_proc = process_data(closed_df, dest_map, lane_map)
        
        # Tabs for Dashboards
        tab1, tab2, tab3 = st.tabs(["🚚 Connection Dashboard", "🏭 Processing Dashboard", "🔍 Bag Search Engine"])
        
        # TAB 1: CONNECTION
        with tab1:
            st.subheader("Line-Haul / Outstation Bags")
            in_conn = in_proc[in_proc['Classification'] == 'Connection']
            cl_conn = cl_proc[cl_proc['Classification'] == 'Connection']
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Connection Bags", len(in_conn) + len(cl_conn))
            col2.metric("Total Shipments", int(in_conn['totalshipments'].sum() + cl_conn['totalshipments'].sum()))
            col3.metric("Critical >24 Hrs", len(in_conn[in_conn['pendinghrs']>24]) + len(cl_conn[cl_conn['pendinghrs']>24]))
            col4.metric("Unmapped Bags", len(in_conn[in_conn['Mapped_Lane']=='Unmapped']) + len(cl_conn[cl_conn['Mapped_Lane']=='Unmapped']))
            
            st.markdown("#### Top 30 High Ageing (Connection)")
            top_conn = pd.concat([in_conn, cl_conn]).sort_values(by='pendinghrs', ascending=False).head(30)
            st.dataframe(top_conn[['bagno', 'bagstatus', 'Clean_Dest', 'Mapped_Lane', 'pendinghrs', 'totalshipments']], use_container_width=True)

        # TAB 2: PROCESSING
        with tab2:
            st.subheader("Local Surat Processing Bags")
            in_pr = in_proc[in_proc['Classification'] == 'Processing']
            cl_pr = cl_proc[cl_proc['Classification'] == 'Processing']
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Processing Bags", len(in_pr) + len(cl_pr))
            c2.metric("Total Shipments", int(in_pr['totalshipments'].sum() + cl_pr['totalshipments'].sum()))
            c3.metric("RTO Bags", len(in_pr[in_pr.get('bagtype','') == 'RTOBag']) + len(cl_pr[cl_pr.get('bagtype','') == 'RTOBag']))
            c4.metric("Avg Pending Hrs", round(pd.concat([in_pr, cl_pr])['pendinghrs'].mean(), 1))
            
            st.markdown("#### Top 30 High Ageing (Processing)")
            top_pr = pd.concat([in_pr, cl_pr]).sort_values(by='pendinghrs', ascending=False).head(30)
            st.dataframe(top_pr[['bagno', 'bagstatus', 'Clean_Dest', 'Mapped_Group', 'pendinghrs', 'totalshipments']], use_container_width=True)

        # TAB 3: SEARCH
        with tab3:
            st.subheader("Bag Search Engine")
            search_bag = st.text_input("Enter Bag Number to Search:")
            if search_bag:
                st.markdown(f"**Results for {search_bag}:**")
                res_in = in_proc[in_proc['bagno'].astype(str).str.contains(search_bag, case=False, na=False)]
                res_cl = cl_proc[cl_proc['bagno'].astype(str).str.contains(search_bag, case=False, na=False)]
                
                if not res_in.empty:
                    st.success("Found in INSCAN!")
                    st.dataframe(res_in[['bagno', 'Clean_Dest', 'Mapped_Lane', 'Classification', 'pendinghrs', 'totalshipments', 'physicalweight', 'lastmodifiedby']])
                if not res_cl.empty:
                    st.warning("Found in BAGCLOSED!")
                    st.dataframe(res_cl[['bagno', 'Clean_Dest', 'Mapped_Lane', 'Classification', 'pendinghrs', 'totalshipments', 'physicalweight', 'lastmodifiedby']])
                if res_in.empty and res_cl.empty:
                    st.error("Bag Not Found in Current Data.")

else:
    st.info("👈 Please upload the Mappings, InScan, and BagClosed files in the Sidebar to generate the dashboard.")
