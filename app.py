import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.14")
st.title("🌐 AICS 北美部署決策中心 (V8.14 最終穩定版)")

# 座標映射表
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    f_df = df.copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        
        mode = st.radio("檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"m_{dimension}")
        chart_type = st.radio("圖表類型", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=f"c_{dimension}")
        sort = st.selectbox("排序", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}")

        # 基礎分組
        df_g = data.groupby(['Month-Date', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
        
        # 繪圖
        x_col = 'Month-Date' if mode == "月份推移" else dimension
        if chart_type == "推移圖": fig = px.line(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension)
        elif chart_type == "柱狀圖": fig = px.bar(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension)
        else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
        st.plotly_chart(fig, use_container_width=True)

        # 數據列表聯動
        if st.checkbox(f"顯示數據列表", key=f"ch_{dimension}"):
            pv = data.pivot_table(index=dimension, columns='Month-Date' if mode=="月份推移" else None, values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
            
            # 計算行列合計
            pv['合計'] = pv.sum(axis=1)
            row_totals = pv.sum(axis=0)
            row_totals.name = '合計'
            
            # 排序 (僅針對數據主體)
            if sort != "預設":
                pv = pv.sort_values(by='合計', ascending=(sort=="由小至大"))
            
            pv = pd.concat([pv, row_totals.to_frame().T])
            st.dataframe(pv)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)
