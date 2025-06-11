import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
import io
import json

# Page configuration
st.set_page_config(
    page_title="Sales Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        border-left: 4px solid;
        margin-bottom: 1rem;
    }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-align: center;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    
    .info-box {
        padding: 1rem 1.5rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    
    .success-box {
        background: #e8f5e8;
        border-left: 4px solid #4CAF50;
        color: #2e7d2e;
    }
    
    .info-box-blue {
        background: #f8f9ff;
        border-left: 4px solid #667eea;
        color: #4a5568;
    }
    
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 16px;
        font-weight: 600;
    }
    
    /* Ranking styles */
    .rank-1 { 
        background: linear-gradient(45deg, #ffd700, #ffed4a) !important; 
        font-weight: bold !important;
    }
    .rank-2 { 
        background: linear-gradient(45deg, #c0c0c0, #e5e5e5) !important; 
        font-weight: bold !important;
    }
    .rank-3 { 
        background: linear-gradient(45deg, #cd7f32, #daa520) !important; 
        font-weight: bold !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Google Sheets Configuration
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def connect_to_google_sheets(sheet_url=None, credentials_json=None):
    """Connect to Google Sheets using service account credentials"""
    try:
        if credentials_json:
            # Use uploaded credentials
            credentials_info = json.loads(credentials_json)
            credentials = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
        else:
            # Use Streamlit secrets (for deployed version)
            credentials = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=SCOPES
            )
        
        client = gspread.authorize(credentials)
        
        if sheet_url:
            # Extract sheet ID from URL
            sheet_id = sheet_url.split('/')[5]
            sheet = client.open_by_key(sheet_id)
            return sheet.sheet1  # First worksheet
        
        return client
    except Exception as e:
        st.error(f"Error connecting to Google Sheets: {str(e)}")
        return None

def load_data_from_google_sheets(worksheet):
    """Load data from Google Sheets"""
    try:
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Error loading data from Google Sheets: {str(e)}")
        return None

def convert_excel_to_sales_data(df):
    """Convert Excel DataFrame to our sales data format"""
    sales_data = []
    
    for _, row in df.iterrows():
        # Handle missing or invalid data
        def safe_get(key, default=0):
            val = row.get(key, default)
            if pd.isna(val) or val == '-' or val == '':
                return default if key.endswith('Appts') else None if 'Close' in key else default
            try:
                return float(val) if 'Close' in key or 'Capture' in key else int(val)
            except:
                return default if key.endswith('Appts') else None if 'Close' in key else default
        
        rep_data = {
            'name': str(row.get('Sales Rep', 'Unknown')),
            'totalAppts': safe_get('Issued Appts', 0),
            'overallClose': safe_get('Overall Close %', 0) * 100 if safe_get('Overall Close %', 0) <= 1 else safe_get('Overall Close %', 0),
            'overallCapture': safe_get('Units Captured on Sold Jobs %', 0) * 100 if safe_get('Units Captured on Sold Jobs %', 0) <= 1 else safe_get('Units Captured on Sold Jobs %', 0),
            'categories': {}
        }
        
        # Process unit categories
        categories = ['0-4', '5-9', '10-17', '18-25', '26+']
        for cat in categories:
            appts_key = f'({cat}) Issued Appts'
            close_key = f'({cat}) Overall Close %'
            capture_key = f'({cat}) Units Captured on Sold Jobs %'
            
            appts = safe_get(appts_key, 0)
            close_rate = safe_get(close_key)
            capture_rate = safe_get(capture_key)
            
            # Convert percentages if they're in decimal format
            if close_rate is not None and close_rate <= 1:
                close_rate *= 100
            if capture_rate is not None and capture_rate <= 1:
                capture_rate *= 100
                
            rep_data['categories'][cat] = {
                'appointments': appts,
                'closeRate': close_rate,
                'captureRate': capture_rate
            }
        
        sales_data.append(rep_data)
    
    return sales_data

def calculate_performance_score(rep):
    """Calculate weighted performance score - identical to React version"""
    unit_categories = ['0-4', '5-9', '10-17', '18-25', '26+']
    valid_categories = []
    
    for cat_key in unit_categories:
        cat_data = rep['categories'][cat_key]
        if (cat_data['appointments'] >= 2 and 
            cat_data['closeRate'] is not None and 
            cat_data['closeRate'] > 0):
            valid_categories.append({
                'appointments': cat_data['appointments'],
                'closeRate': cat_data['closeRate'],
                'captureRate': cat_data['captureRate'] or 0
            })
    
    # Calculate weighted average for categories
    total_weighted_close = 0
    total_weighted_capture = 0
    total_weight = 0
    
    for cat in valid_categories:
        weight = cat['appointments']
        total_weighted_close += cat['closeRate'] * weight
        total_weighted_capture += cat['captureRate'] * weight
        total_weight += weight
    
    avg_category_close = total_weighted_close / total_weight if total_weight > 0 else 0
    avg_category_capture = total_weighted_capture / total_weight if total_weight > 0 else 0
    
    # Normalize capture rate
    normalized_capture = min(avg_category_capture, 150)
    
    # Apply weighting: 50% overall close, 15% category close, 35% capture
    score = (rep['overallClose'] * 0.50) + (avg_category_close * 0.15) + (normalized_capture * 0.35)
    
    return {
        'name': rep['name'],
        'score': score,
        'overallClose': rep['overallClose'],
        'avgCategoryClose': avg_category_close,
        'avgCategoryCapture': avg_category_capture,
        'totalAppts': rep['totalAppts'],
        'validCategories': len(valid_categories),
        'rawData': rep
    }

# Initialize session state
if 'sales_data' not in st.session_state:
    # Default sample data (your current data)
    st.session_state.sales_data = [
        {"name": "Gabriel Grimm", "totalAppts": 32, "overallClose": 31.25, "overallCapture": 94.20, 
         "categories": {"0-4": {"appointments": 9, "closeRate": 22.22, "captureRate": 222.54}, 
                       "5-9": {"appointments": 2, "closeRate": 0, "captureRate": 119.78}, 
                       "10-17": {"appointments": 10, "closeRate": 30, "captureRate": 97.43}, 
                       "18-25": {"appointments": 6, "closeRate": 66.67, "captureRate": 102.35}, 
                       "26+": {"appointments": 5, "closeRate": 20, "captureRate": 52.78}}},
        {"name": "Derek Kingry", "totalAppts": 39, "overallClose": 38.46, "overallCapture": 107.17, 
         "categories": {"0-4": {"appointments": 15, "closeRate": 40, "captureRate": 133.59}, 
                       "5-9": {"appointments": 9, "closeRate": 44.44, "captureRate": 122.89}, 
                       "10-17": {"appointments": 8, "closeRate": 12.5, "captureRate": 100.67}, 
                       "18-25": {"appointments": 7, "closeRate": 57.14, "captureRate": 66.67}, 
                       "26+": {"appointments": 0, "closeRate": None, "captureRate": None}}},
        {"name": "Craig Chisman", "totalAppts": 29, "overallClose": 44.83, "overallCapture": 98.0, 
         "categories": {"0-4": {"appointments": 14, "closeRate": 50, "captureRate": 149.0}, 
                       "5-9": {"appointments": 7, "closeRate": 42.86, "captureRate": 120.0}, 
                       "10-17": {"appointments": 5, "closeRate": 40, "captureRate": 71.0}, 
                       "18-25": {"appointments": 2, "closeRate": 50, "captureRate": 60.0}, 
                       "26+": {"appointments": 1, "closeRate": 0, "captureRate": None}}}
    ]

# Calculate rankings
ranked_reps = sorted([calculate_performance_score(rep) for rep in st.session_state.sales_data], 
                    key=lambda x: x['score'], reverse=True)

# HEADER
st.markdown('<h1 class="main-header">Sales Rep Performance Rankings</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; font-size: 1.2rem; color: #666; margin-bottom: 2rem;">Complete Unit Category Analysis - All 5 Categories</p>', unsafe_allow_html=True)

# SIDEBAR - Data Management
st.sidebar.header("📊 Data Management")

data_source = st.sidebar.radio(
    "Choose Data Source:",
    ["Current Data", "Upload Excel File", "Google Sheets", "Sample Data"]
)

if data_source == "Upload Excel File":
    st.sidebar.subheader("📁 Upload Excel File")
    uploaded_file = st.sidebar.file_uploader(
        "Choose Excel file", 
        type=['xlsx', 'xls'],
        help="Upload your weekly sales performance Excel file"
    )
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            new_sales_data = convert_excel_to_sales_data(df)
            st.session_state.sales_data = new_sales_data
            ranked_reps = sorted([calculate_performance_score(rep) for rep in st.session_state.sales_data], 
                               key=lambda x: x['score'], reverse=True)
            st.sidebar.success(f"✅ Loaded {len(new_sales_data)} sales reps")
            
            with st.sidebar.expander("Preview Data"):
                st.dataframe(df.head())
                
        except Exception as e:
            st.sidebar.error(f"❌ Error reading file: {str(e)}")

elif data_source == "Google Sheets":
    st.sidebar.subheader("📊 Google Sheets Integration")
    
    # Option 1: Public sheet URL
    sheet_url = st.sidebar.text_input(
        "Google Sheets URL:",
        placeholder="https://docs.google.com/spreadsheets/d/...",
        help="Make sure the sheet is publicly viewable"
    )
    
    # Option 2: Service account credentials
    credentials_file = st.sidebar.file_uploader(
        "Service Account JSON (Optional):",
        type=['json'],
        help="Upload Google service account credentials for private sheets"
    )
    
    if st.sidebar.button("Load from Google Sheets"):
        if sheet_url:
            try:
                credentials_json = None
                if credentials_file:
                    credentials_json = credentials_file.read().decode('utf-8')
                
                worksheet = connect_to_google_sheets(sheet_url, credentials_json)
                if worksheet:
                    df = load_data_from_google_sheets(worksheet)
                    if df is not None:
                        new_sales_data = convert_excel_to_sales_data(df)
                        st.session_state.sales_data = new_sales_data
                        ranked_reps = sorted([calculate_performance_score(rep) for rep in st.session_state.sales_data], 
                                           key=lambda x: x['score'], reverse=True)
                        st.sidebar.success(f"✅ Loaded {len(new_sales_data)} reps from Google Sheets")
            except Exception as e:
                st.sidebar.error(f"❌ Error: {str(e)}")

# INFO BOXES
col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    <div class="info-box success-box">
        <strong>📈 Complete Performance Data:</strong> Analyzing sales reps across all 5 unit categories (0-4, 5-9, 10-17, 18-25, 26+ units) • Updated weighting: 50% Overall Close + 35% Capture + 15% Category Close
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="info-box info-box-blue">
        <strong>Scoring Method:</strong> Weighted performance score using <strong>50% Overall Close Rate + 15% Category Close Rate + 35% Capture Rate</strong> • All unit sizes from Small (0-4) to Mega Jobs (26+)
    </div>
    """, unsafe_allow_html=True)

# QUICK STATS
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="👥 Total Reps",
        value=len(st.session_state.sales_data),
        help="Active sales representatives"
    )

with col2:
    avg_close = sum(rep['overallClose'] for rep in st.session_state.sales_data) / len(st.session_state.sales_data)
    st.metric(
        label="🎯 Avg Close Rate",
        value=f"{avg_close:.1f}%",
        help="Overall performance"
    )

with col3:
    avg_capture = sum(min(rep['overallCapture'], 150) for rep in st.session_state.sales_data) / len(st.session_state.sales_data)
    st.metric(
        label="📈 Avg Capture Rate", 
        value=f"{avg_capture:.1f}%",
        help="Unit capture efficiency"
    )

with col4:
    top_performer = ranked_reps[0] if ranked_reps else {'name': 'N/A', 'score': 0}
    st.metric(
        label="🏆 Top Performer",
        value=top_performer['name'].split()[0],
        delta=f"Score: {top_performer['score']:.1f}",
        help="Highest weighted score"
    )

# MAIN TABS
tab1, tab2, tab3 = st.tabs(["📊 Overall Rankings", "🏆 Category Leaders", "📋 Detailed Analysis"])

with tab1:
    st.subheader("Top Performers - Overall Weighted Score")
    
    # Filter
    col1, col2 = st.columns([3, 1])
    with col1:
        min_appointments = st.selectbox(
            "Minimum Appointments:",
            options=[0, 10, 20, 25],
            index=1,
            format_func=lambda x: f"{x}+ Appointments" if x > 0 else "All Reps"
        )
    
    # Filter data
    filtered_reps = [rep for rep in ranked_reps if rep['totalAppts'] >= min_appointments]
    
    st.success("✅ **Updated with Latest Data:** New scoring system (50% Overall Close + 35% Capture + 15% Category Close) with all 5 unit categories including 26+ Mega Jobs.")
    
    # Create rankings table
    if filtered_reps:
        rankings_data = []
        for i, rep in enumerate(filtered_reps):
            rank_emoji = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else "📍"
            rankings_data.append({
                'Rank': f"{rank_emoji} {i+1}",
                'Sales Rep': rep['name'],
                'Score': f"{rep['score']:.1f}",
                'Overall Close %': f"{rep['overallClose']:.1f}%",
                'Avg Category Close %': f"{rep['avgCategoryClose']:.1f}%", 
                'Avg Capture %': f"{rep['avgCategoryCapture']:.0f}%",
                'Total Appts': rep['totalAppts'],
                'Active Categories': rep['validCategories']
            })
        
        df_rankings = pd.DataFrame(rankings_data)
        
        # Display with styling
        st.dataframe(
            df_rankings,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("No reps meet the minimum appointment criteria.")

with tab2:
    st.subheader("Category Leaders")
    
    # Category definitions
    unit_categories = ['0-4', '5-9', '10-17', '18-25', '26+']
    category_names = {
        '0-4': 'Small Jobs (0-4 Units)',
        '5-9': 'Medium Jobs (5-9 Units)', 
        '10-17': 'Large Jobs (10-17 Units)',
        '18-25': 'Extra Large Jobs (18-25 Units)',
        '26+': 'Mega Jobs (26+ Units)'
    }
    
    # Calculate category leaders
    category_leaders = {}
    for cat_key in unit_categories:
        category_reps = []
        for rep in st.session_state.sales_data:
            cat_data = rep['categories'][cat_key]
            if (cat_data['appointments'] >= 3 and 
                cat_data['closeRate'] is not None and 
                cat_data['closeRate'] > 0):
                category_score = (cat_data['closeRate'] * 0.6) + (min(cat_data['captureRate'] or 0, 150) * 0.4)
                category_reps.append((rep, category_score))
        
        if category_reps:
            category_leaders[cat_key] = max(category_reps, key=lambda x: x[1])[0]
    
    # Leader cards
    cols = st.columns(5)
    for i, cat_key in enumerate(unit_categories):
        with cols[i]:
            st.markdown(f"#### {category_names[cat_key]}")
            if cat_key in category_leaders:
                leader = category_leaders[cat_key]
                cat_data = leader['categories'][cat_key]
                st.success(f"🏆 **{leader['name']}**")
                st.write(f"📊 {cat_data['closeRate']:.1f}% Close Rate")
                st.write(f"📈 {cat_data['captureRate']:.0f}% Capture Rate")
                st.write(f"📅 {cat_data['appointments']} Appointments")
            else:
                st.warning("Building Data")
                st.write("Need 3+ appointments for reliable ranking")
    
    st.markdown("---")
    
    # Category ranking tables
    for cat_key in unit_categories:
        with st.expander(f"📊 {category_names[cat_key]} - Detailed Rankings", expanded=False):
            category_reps = []
            for rep in st.session_state.sales_data:
                cat_data = rep['categories'][cat_key]
                if (cat_data['appointments'] >= 2 and 
                    cat_data['closeRate'] is not None and 
                    cat_data['closeRate'] > 0):
                    category_score = (cat_data['closeRate'] * 0.6) + (min(cat_data['captureRate'] or 0, 150) * 0.4)
                    category_reps.append({
                        'Rank': len(category_reps) + 1,
                        'Sales Rep': rep['name'],
                        'Close Rate': f"{cat_data['closeRate']:.1f}%",
                        'Capture Rate': f"{cat_data['captureRate']:.0f}%",
                        'Appointments': cat_data['appointments'],
                        'Score': category_score
                    })
            
            if category_reps:
                df_cat = pd.DataFrame(sorted(category_reps, key=lambda x: x['Score'], reverse=True)[:10])
                df_cat['Rank'] = range(1, len(df_cat) + 1)
                df_cat = df_cat.drop('Score', axis=1)
                st.dataframe(df_cat, hide_index=True, use_container_width=True)
            else:
                st.info("No reps with sufficient data in this category")

with tab3:
    st.subheader("Complete Performance Matrix")
    st.caption("Format: Close Rate / Capture Rate (Appointments)")
    
    # Create detailed matrix
    if ranked_reps:
        matrix_data = []
        for rep in ranked_reps:
            row_data = {'Sales Rep': rep['name'], 'Overall Score': f"{rep['score']:.1f}"}
            
            for cat_key in unit_categories:
                cat_data = rep['rawData']['categories'][cat_key]
                if (cat_data['appointments'] >= 2 and cat_data['closeRate'] is not None):
                    row_data[f"{cat_key} Units"] = f"{cat_data['closeRate']:.1f}% / {cat_data['captureRate']:.0f}% ({cat_data['appointments']})"
                else:
                    row_data[f"{cat_key} Units"] = "—"
            
            row_data['Total Appts'] = rep['totalAppts']
            matrix_data.append(row_data)
        
        df_matrix = pd.DataFrame(matrix_data)
        st.dataframe(df_matrix, hide_index=True, use_container_width=True)

# FOOTER
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p><strong>Sales Performance Dashboard</strong> | Identical calculations to React version | Weekly data updates supported</p>
    <p>Data Source: {data_source} | Last Updated: {timestamp}</p>
</div>
""".format(
    data_source=data_source,
    timestamp=pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
), unsafe_allow_html=True)