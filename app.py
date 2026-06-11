import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from fpdf import FPDF
import io
import datetime

# 1. 初始化頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.4")
st.title("🌐 AICS 北美部署決策中心 (V6.4 多視角穿透分析版)")

# CSS 樣式注入
st.markdown("""
    <style>
    .big-metric { font-size: 24px; font-weight: bold; color: #e63946; }
    .label-text { font-size: 14px; color: #457b9d; }
    </style>
    """, unsafe_allow_html=True)

# 數據導入
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

    # 封裝分析模組 (已新增視角切換邏輯)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        
        # 模式切換：月份趨勢 vs 全量彙總
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

    # 執行各維度渲染
    dims = [
        ('Machine Type', '設備維度 (Machine Type)'),
        ('Field', '場域維度 (Field)'),
        ('Area', '區域維度 (Area)'),
        ('Company', '客戶維度 (Company)'),
        ('Device/Platform', '平台維度 (Device/Platform)')
    ]
    
    for dim, name in dims:
        render_analysis_section(f_df, dim, name)

    # 匯出功能
    if st.sidebar.button("📊 導出完整報告"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image:
            slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width)
        buf_ppt = io.BytesIO()
        prs.save(buf_ppt)
        st.sidebar.download_button("下載 PPTX", buf_ppt.getvalue(), "Tactical_Report.pptx")

else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
