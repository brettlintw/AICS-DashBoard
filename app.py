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
st.markdown("""<style>.stDataFrame table td, .stDataFrame table th { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)
st.title("🌐 AICS 北美部署決策中心 (V8.10 完整功能版)")

# 座標映射表
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    st.subheader("📊 設備總覽統計")
    st.dataframe(f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index(), use_container_width=True)

    # 1. 地圖 (修正版)
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    for state, coords in US_STATES_COORDS.items():
        fig_map.add_trace(go.Scattergeo(lon=[coords[1]], lat=[coords[0]], text=[state], mode='text', textfont=dict(size=9, color='darkblue')))
    for m in selected_machines:
        m_df = f_df[f_df['Machine Type'] == m].groupby('State Code')['Outbound Qty (Item)'].sum().reset_index()
        fig_map.add_trace(go.Scattergeo(locations=m_df['State Code'], locationmode="USA-states", marker=dict(size=m_df['Outbound Qty (Item)']*0.8, opacity=0.6), name=m))
    fig_map.update_layout(geo=dict(scope='usa', projection_type='albers usa', projection_scale=1.0, fitbounds="locations"), height=500, margin={"r": 0, "t": 0, "l": 0, "b": 0})
    st.plotly_chart(fig_map, use_container_width=True)

    # 2. 維度分析模組
    def render_analysis_section(dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        c1, c2 = st.columns([1, 4])
        with c1:
            mode = st.radio("檢視模式", ["月份推移", "全時間段"], key=f"m_{dimension}")
            chart_type = st.radio("圖表類型", ["柱狀圖", "推移圖", "餅圖"], key=f"c_{dimension}")
            sort = st.selectbox("排序", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}") if mode == "全時間段" else "預設"
        with c2:
            df_g = f_df.groupby(['Month-Date', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
            if mode == "全時間段" and sort != "預設": df_g = df_g.sort_values(by='Outbound Qty (Item)', ascending=(sort=="由小至大"))
            
            if chart_type == "柱狀圖": fig = px.bar(df_g, x=dimension if mode=="全時間段" else "Month-Date", y='Outbound Qty (Item)', color=dimension, text='Outbound Qty (Item)')
            elif chart_type == "推移圖": fig = px.line(df_g, x="Month-Date", y='Outbound Qty (Item)', color=dimension, markers=True)
            else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df_g, use_container_width=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(dim, name)

    # 3. PPTX 匯出 (維持專業報表)
    if st.sidebar.button("📊 導出完整戰情室報表"):
        prs = Presentation()
        for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.text = f"分析維度: {name}"
            df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
            try:
                fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)')
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
