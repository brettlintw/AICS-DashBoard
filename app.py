import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from fpdf import FPDF
import io
import datetime

# 1. 初始化頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.4-修正版")
st.title("🌐 AICS 北美部署決策中心 (V6.4 修正版：新增全時間段彙總)")

# CSS 樣式
st.markdown("""
    <style>
    .big-metric { font-size: 24px; font-weight: bold; color: #e63946; }
    .label-text { font-size: 14px; color: #457b9d; }
    </style>
    """, unsafe_allow_html=True)

# 座標資料
US_STATES_COORDS = {
    'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6],
    'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6],
    'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2],
    'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8],
    'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2],
    'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5],
    'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7],
    'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9],
    'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7],
    'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]
}

uploaded_file = st.sidebar.file_uploader("上傳 Excel 數據 (Data Base)", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳戰報背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    df['Month-Year'] = df['Date(出庫)'].dt.strftime('%Y-%m')

    all_machines = df['Machine Type'].unique().tolist()
    selected_machines = st.sidebar.multiselect("設備類型篩選", all_machines, default=all_machines)
    f_df = df[df['Machine Type'].isin(selected_machines)].copy()

    # 地圖模組 (補回)
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    map_agg = f_df.groupby(['State Code', 'Machine Type'])['Outbound Qty (Item)'].sum().reset_index()
    for m in selected_machines:
        d = map_agg[map_agg['Machine Type'] == m]
        fig_map.add_trace(go.Scattergeo(locations=d['State Code'], locationmode="USA-states", marker=dict(size=d['Outbound Qty (Item)'], opacity=0.7), name=m))
    fig_map.update_layout(geo=dict(scope='usa'), height=400)
    st.plotly_chart(fig_map, use_container_width=True)

    # 分析模組 (新增全時間段檢視)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        view_mode = st.radio(f"檢視模式 ({title_name})", ["月份推移分析", "全時間段數據彙總"], horizontal=True, key=f"mode_{dimension}")
        
        if view_mode == "月份推移分析":
            df_group = data.groupby(['Month-Year', dimension])['Outbound Qty (Item)'].sum().reset_index()
            fig = px.bar(df_group, x='Month-Year', y='Outbound Qty (Item)', color=dimension, barmode='group')
            pivot = data.pivot_table(index=dimension, columns='Month-Year', values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
        else:
            df_total = data.groupby(dimension)[['Outbound Qty (Item)']].sum().reset_index()
            fig = px.pie(df_total, values='Outbound Qty (Item)', names=dimension, hole=0.3)
            pivot = df_total.set_index(dimension)
            
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(pivot.style.format("{:.0f}"), use_container_width=True)

    # 執行所有維度
    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    if st.sidebar.button("📊 導出完整報告"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image: slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width)
        buf_ppt = io.BytesIO(); prs.save(buf_ppt)
        st.sidebar.download_button("下載 PPTX", buf_ppt.getvalue(), "Tactical_Report.pptx")
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
