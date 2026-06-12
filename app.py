import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import requests

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.14 (雲端+地圖版)")
st.title("🌐 AICS 北美部署決策中心 (雲端地圖版)")

# 座標映射表 (完整保留)
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

# --- 雲端數據讀取模組 ---
SHEET_ID = "https://docs.google.com/spreadsheets/d/1AQNSCVWVoY9dFWD4NdCmXRWiRQt0--LE/edit?usp=sharing&ouid=101935051207867221848&rtpof=true&sd=true" 

@st.cache_data(ttl=600)
def load_data():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=Data+Base"
        response = requests.get(url)
        response.encoding = 'utf-8'
        return pd.read_csv(io.StringIO(response.text))
    except Exception as e:
        st.error(f"連線錯誤: {e}")
        return None

df = load_data()
if df is not None:
    df.columns = df.columns.str.strip().str.replace('\u200b', '', regex=False)
    if 'Date(出庫)' in df.columns:
        df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
        f_df = df.copy()
        f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')
    else:
        st.stop()
else:
    st.stop()

# --- 地圖渲染模組 (保留 V8.10 邏輯) ---
st.subheader("📍 北美部署地圖")
map_data = f_df.groupby('State Code')[['Outbound Qty (Item)']].sum().reset_index()
map_data['lat'] = map_data['State Code'].map(lambda x: US_STATES_COORDS.get(x, [0, 0])[0])
map_data['lon'] = map_data['State Code'].map(lambda x: US_STATES_COORDS.get(x, [0, 0])[1])

fig_map = px.scatter_geo(map_data, lat='lat', lon='lon', size='Outbound Qty (Item)', 
                         scope='usa', hover_name='State Code', title="各州部署密度")
st.plotly_chart(fig_map, use_container_width=True)

# --- 核心分析功能模組 (保留 V8.10 邏輯) ---
def render_analysis_section(data, dimension, title_name):
    st.markdown("---")
    st.subheader(f"📈 {title_name}")
    col1, col2 = st.columns([1, 1])
    with col1:
        mode = st.radio("檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"m_{dimension}")
        chart_type = st.radio("圖表類型", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=f"c_{dimension}")
    with col2:
        sort = st.selectbox("排序", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}")

    df_g = data.groupby(['Month-Date', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
    x_col = 'Month-Date' if mode == "月份推移" else dimension
    
    if chart_type == "推移圖": fig = px.line(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension)
    elif chart_type == "柱狀圖": fig = px.bar(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension)
    else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
    st.plotly_chart(fig, use_container_width=True)

    if st.checkbox(f"顯示數據列表", key=f"ch_{dimension}"):
        pivot = data.pivot_table(index=dimension, columns='Month-Date' if mode=="月份推移" else None, values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
        pivot['合計'] = pivot.sum(axis=1)
        if sort != "預設": pivot = pivot.sort_values(by='合計', ascending=(sort=="由小至大"))
        col_totals = pivot.sum(axis=0); col_totals.name = '合計'
        pivot = pd.concat([pivot, col_totals.to_frame().T])
        def highlight_row(row): return ['color: blue; font-weight: bold;' if row.name == '合計' else ''] * len(row)
        st.dataframe(pivot.style.apply(highlight_row, axis=1), use_container_width=True)

for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
    render_analysis_section(f_df, dim, name)
