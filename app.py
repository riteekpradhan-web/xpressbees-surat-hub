import streamlit as st
import pandas as pd
import io

# --- 1. PAGE CONFIG (Sabse upar hona zaroori hai) ---
st.set_page_config(page_title="XpressBees Surat Dashboard", layout="wide")

# --- 2. SECURITY GATE ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 XpressBees Secure Portal")
    pwd = st.text_input("Enter Hub Password:", type="password")
    if st.button("Login"):
        if pwd == st.secrets.get("APP_PASSWORD", "SuratHub@2026"): # Fallback for safety
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ Galat Password! Access Denied.")
    st.stop()

# --- 3. MAIN DASHBOARD UI ---
st.title("👑 XpressBees Surat Hub - Deep Analysis Dashboard")
st.markdown("**Owner: Pandi Jagan** | Live Operations Portal")

st.sidebar.header("📂 Data Upload (Daily)")
mapping_file = st.sidebar.file_uploader("1. Upload Mappings (Expriment 2.xlsx)", type=['xlsx'])
inscan_file = st.sidebar.file_uploader("2. Upload InScan CSV", type=['csv'])
closed_file = st.sidebar.file_uploader("3. Upload BagClosed CSV", type=['csv'])

# Function to convert dataframe to excel for download
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

@st.cache_data
def load_mappings(file):
    try:
        xls = pd.ExcelFile(file)
        dest_df = pd.read_excel(xls, "GroupName").fillna("")
        lane_df = pd.read_excel(xls, "lANE").fillna("")
        dest_map = dict(zip(dest_df['ShipmentDestination'].astype(str).str.strip().str.upper(), dest_df['GroupName'].astype(str).str.strip()))
        lane_map = dict(zip(lane_df['Groups'].astype(str).str.strip().str.upper(), lane_df['lANE'].astype(str).str.strip()))
        return dest_map, lane_map
    except:
        return {}, {}

if mapping_file and inscan_file and closed_file:
    dest_map, lane_map = load_mappings(mapping_file)
    
    in_df = pd.read_csv(inscan_file)
    cl_df = pd.read_csv(closed_file)
    
    def process_data(df, source):
        df = df.copy()
        df['Source'] = source
        # Drop duplicate bags in the same file to fix ghost counting
        df = df.drop_duplicates(subset=['bagno'], keep='first')
        
        df['totalshipments'] = pd.to_numeric(df.get('totalshipments', 0), errors='coerce').fillna(0)
        df['physicalweight'] = pd.to_numeric(df.get('physicalweight', 0), errors='coerce').fillna(0)
        df['pendinghrs'] = pd.to_numeric(df.get('pendinghrs', 0), errors='coerce').fillna(0)
        
        df['Clean_Dest'] = df.get('destinationhubname', '').astype(str).apply(lambda x: x.split('(')[0].strip().upper())
        df['Mapped_Group'] = df['Clean_Dest'].apply(lambda x: dest_map.get(x, "UNMAPPED"))
        df['Mapped_Lane'] = df['Mapped_Group'].apply(lambda x: lane_map.get(x.upper(), "UNMAPPED") if x != "UNMAPPED" else "UNMAPPED")
        
        # Classification (Processing vs Connection)
        df['Classification'] = df.apply(lambda r: "Processing" if r['Clean_Dest'].startswith("SRT") or r['Mapped_Group'] == "SURAT CITY" or r['Mapped_Lane'] == "LOCAL" else "Connection", axis=1)
        
        # Exact Ageing Buckets
        df['Age_Bucket'] = pd.cut(df['pendinghrs'], bins=[-1, 1, 3, 12, 24, 9999], labels=['0-1 hr', '1-3 hrs', '3-12 hrs', '12-24 hrs', '>24 hrs'])
        return df

    in_proc = process_data(in_df, "InScan")
    cl_proc = process_data(cl_df, "BagClosed")
    all_data = pd.concat([in_proc, cl_proc])

    tab1, tab2, tab3 = st.tabs(["🚚 Connection Dashboard", "🏭 Processing Dashboard", "🔍 Bag Search Engine"])
    
    with tab1:
        st.subheader("Lane-Wise Connection Summary")
        conn_data = all_data[all_data['Classification'] == 'Connection']
        if not conn_data.empty:
            # Pivot table to merge duplicate lanes into a single row
            pivot_conn = conn_data.pivot_table(
                index='Mapped_Lane', 
                values=['bagno', 'totalshipments', 'physicalweight'],
                aggfunc={'bagno': 'count', 'totalshipments': 'sum', 'physicalweight': 'sum'}
            ).rename(columns={'bagno': 'Total Bags', 'totalshipments': 'Total Shipments', 'physicalweight': 'Total Weight'})
            
            ageing_conn = conn_data.pivot_table(index='Mapped_Lane', columns='Age_Bucket', values='bagno', aggfunc='count', fill_value=0)
            final_conn = pivot_conn.join(ageing_conn).reset_index()
            
            st.dataframe(final_conn, use_container_width=True)
            
            # Download Actionable Bags Button
            st.download_button(label="📥 Download Connection Bags Details", data=to_excel(conn_data), file_name="Connection_Bags_Details.xlsx", mime="application/vnd.ms-excel")

    with tab2:
        st.subheader("Hub-Wise Processing Summary")
        proc_data = all_data[all_data['Classification'] == 'Processing']
        if not proc_data.empty:
            pivot_proc = proc_data.pivot_table(
                index='Clean_Dest', 
                values=['bagno', 'totalshipments', 'physicalweight'],
                aggfunc={'bagno': 'count', 'totalshipments': 'sum', 'physicalweight': 'sum'}
            ).rename(columns={'bagno': 'Total Bags', 'totalshipments': 'Total Shipments', 'physicalweight': 'Total Weight'})
            
            ageing_proc = proc_data.pivot_table(index='Clean_Dest', columns='Age_Bucket', values='bagno', aggfunc='count', fill_value=0)
            final_proc = pivot_proc.join(ageing_proc).reset_index()
            
            st.dataframe(final_proc, use_container_width=True)
            
            # Download Actionable Bags Button
            st.download_button(label="📥 Download Processing Bags Details", data=to_excel(proc_data), file_name="Processing_Bags_Details.xlsx", mime="application/vnd.ms-excel")

    with tab3:
        st.subheader("Bag Search Engine")
        search_val = st.text_input("Enter Bag Number:")
        if search_val:
            res = all_data[all_data['bagno'].astype(str).str.contains(search_val, case=False, na=False)]
            if not res.empty:
                st.dataframe(res[['Source', 'bagno', 'Clean_Dest', 'Mapped_Group', 'Mapped_Lane', 'Classification', 'pendinghrs', 'totalshipments', 'bagcreatedby', 'lastmodifiedby']], use_container_width=True)
            else:
                st.error("Bag Not Found!")
