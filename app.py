import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
import plotly.io as pio

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.10")
st.title("🌐 AICS 北美部署決策中心 (V8.10 動態圖表版)")

# [數據導入區與地圖邏輯保持不變...]
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # UI：各維度圖表 (包含切換功能)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        col1, col2 = st.columns([1, 3])
        with col1:
            chart_type = st.radio("圖表類型", ["柱狀圖", "推移圖", "餅圖"], key=f"c_{dimension}")
        with col2:
            df_g = data.groupby([dimension])[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
            if chart_type == "柱狀圖":
                fig = px.bar(df_g, x=dimension, y='Outbound Qty (Item)', text='Outbound Qty (Item)')
            elif chart_type == "推移圖":
                # 推移圖需加入時間軸數據
                df_t = data.groupby(['Month-Date', dimension])[['Outbound Qty (Item)']].sum().reset_index()
                fig = px.line(df_t, x='Month-Date', y='Outbound Qty (Item)', color=dimension, markers=True)
            else:
                fig = px.pie(df_g, names=dimension, values='Outbound Qty (Item)')
            st.plotly_chart(fig, use_container_width=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    # 4. PPTX 導出 (維持專業報表佈局)
    if st.sidebar.button("📊 導出完整戰情室報表"):
        try:
            prs = Presentation()
            for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.text = f"分析維度: {name}"
                
                df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
                
                # PPT 統一輸出柱狀圖以確保版面專業
                fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)')
                try:
                    img_bytes = pio.to_image(fig, format="png", width=1200, height=450)
                    slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(1), width=Inches(9))
                except: pass
                
                rows, cols = df_g.shape
                table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(5), Inches(9), Inches(2)).table
                for i, col in enumerate(df_g.columns): table.cell(0, i).text = str(col)
                for r in range(rows):
                    for c in range(cols): table.cell(r + 1, c).text = str(df_g.iloc[r, c])
            
            buf = io.BytesIO()
            prs.save(buf)
            st.sidebar.download_button("下載報表", buf.getvalue(), "Tactical_Report.pptx")
            st.sidebar.success("導出完成！")
        except Exception as e:
            st.sidebar.error(f"導出異常: {e}")
