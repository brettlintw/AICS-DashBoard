import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io

# 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.9.1")

# CSS 樣式：指定藍色與黑色的呈現規則
st.markdown("""
    <style>
    .blue-text { color: #0000ff !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# 顏色標示函數：僅針對索引名為『合計』的列標示為藍色
def highlight_total(row):
    color = '#0000ff' if row.name == '合計' or '合計' in str(row.name) else 'black'
    return [f'color: {color}; font-weight: bold;' if color == '#0000ff' else 'color: black' for _ in row]

st.title("🌐 AICS 北美部署決策中心 (V6.9.1 色彩優化版)")

# 側邊欄控制台
st.sidebar.header("⚙️ 戰情控制台")
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Year'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # 1. 設備總覽統計 (應用色彩函數)
    st.subheader("📊 設備總覽統計")
    summary = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    summary.loc[len(summary)] = ['合計', summary['Outbound Qty (Item)'].sum()]
    summary = summary.set_index('Machine Type')
    st.dataframe(summary.style.apply(highlight_total, axis=1), use_container_width=True)

    # 2. 北美地圖
    st.subheader("🗺️ 北美設備戰術分佈")
    # ... (地圖邏輯沿用) ...
    
    # 3. 模組化分析 (應用色彩函數)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        mode = st.radio("檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"m_{dimension}")
        sort = st.selectbox("排序方式", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}") if mode == "全時間段彙總" else "預設"

        df_g = data.groupby(['Month-Year', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
        
        if sort == "由大至小": df_g = df_g.sort_values(by='Outbound Qty (Item)', ascending=False)
        elif sort == "由小至大": df_g = df_g.sort_values(by='Outbound Qty (Item)', ascending=True)

        # 繪圖邏輯...
        
        if st.checkbox(f"顯示數據列表", key=f"ch_{dimension}"):
            pivot = data.pivot_table(index=dimension, columns='Month-Year' if mode=="月份推移" else None, values='Outbound Qty (Item)', aggfunc='sum', fill_value=0, margins=True, margins_name='合計')
            st.dataframe(pivot.style.apply(highlight_total, axis=1), use_container_width=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    # 匯出邏輯...
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
