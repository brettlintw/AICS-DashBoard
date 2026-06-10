import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from fpdf import FPDF
import io
import datetime

# 1. 初始化頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.3")
st.title("🌐 AICS 北美部署決策中心 (V6.3 全維度整合版)")

# CSS 樣式注入
st.markdown("""
    <style>
    .big-metric { font-size: 24px; font-weight: bold; color: #e63946; }
    .label-text { font-size: 14px; color: #457b9d; }
    </style>
    """, unsafe_allow_html=True)

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

# 2. 數據導入與處理
st.sidebar.header("📥 數據導入與匯出")
uploaded_file = st.sidebar.file_uploader("上傳 Excel 數據 (需包含 Data Base 分頁)", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳 PPT 戰報背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    df['Month-Year'] = df['Date(出庫)'].dt.strftime('%Y-%m')

    machine_list = df['Machine Type'].unique().tolist()
    selected_machines = st.sidebar.multiselect("設備類型選擇", machine_list, default=machine_list)
    date_range = st.sidebar.date_input("日期區間", value=(df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()))
    
    f_df = df[(df['Machine Type'].isin(selected_machines)) & 
              (df['Date(出庫)'].dt.date >= date_range[0]) & 
              (df['Date(出庫)'].dt.date <= date_range[1])].copy()

    # 3. 指標顯示
    st.markdown("<br><br>", unsafe_allow_html=True)
    cols = st.columns(len(selected_machines) if len(selected_machines) > 0 else 1)
    for i, m in enumerate(selected_machines):
        val = int(f_df[f_df['Machine Type'] == m]['Outbound Qty (Item)'].sum())
        cols[i].markdown(f"<div class='label-text'>{m}</div><div class='big-metric'>{val} 台</div>", unsafe_allow_html=True)
    st.markdown("<br><br>", unsafe_allow_html=True)

        # 4. 北美分布地圖
    st.subheader("🗺️ 北美設備戰術分佈 (全州代碼標註)")
    fig_map = go.Figure()
    
    fig_map.add_trace(go.Scattergeo(
        lon=[US_STATES_COORDS[s][1] for s in US_STATES_COORDS],
        lat=[US_STATES_COORDS[s][0] for s in US_STATES_COORDS],
        text=list(US_STATES_COORDS.keys()),
        mode='text',
        textfont=dict(size=14, color="black", family="Arial Black"),
        showlegend=False
    ))
    
    if not f_df.empty:
        map_agg = f_df.groupby(['State Code', 'Machine Type'])['Outbound Qty (Item)'].sum().reset_index()
        for machine in selected_machines:
            plot_data = map_agg[map_agg['Machine Type'] == machine]
            fig_map.add_trace(go.Scattergeo(
                locations=plot_data['State Code'], locationmode="USA-states",
                marker=dict(size=plot_data['Outbound Qty (Item)']*3, opacity=0.6, line_width=0),
                name=machine, text=plot_data['Outbound Qty (Item)'], hovertemplate="%{location}: %{text}台"
            ))
    
    fig_map.update_layout(geo=dict(scope='usa'), height=600, margin={"r":0,"t":0,"l":0,"b":0})
    st.plotly_chart(fig_map, use_container_width=True)

    # 5. 分析模組
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name} 分析")
        chart_type = st.radio(f"選擇 {title_name} 圖表", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=title_name)
        df_group = data.groupby(['Month-Year', dimension])['Outbound Qty (Item)'].sum().reset_index().sort_values('Month-Year')
        
        if chart_type == "推移圖":
            fig = px.line(df_group, x='Month-Year', y='Outbound Qty (Item)', color=dimension, markers=True, text='Outbound Qty (Item)')
            fig.update_traces(line=dict(width=3, shape='spline'), textposition="top center")
        elif chart_type == "柱狀圖":
            fig = px.bar(df_group, x='Month-Year', y='Outbound Qty (Item)', color=dimension, barmode='group', text='Outbound Qty (Item)')
            fig.update_traces(texttemplate='%{text}', textposition='outside')
        else:
            fig = px.pie(data, values='Outbound Qty (Item)', names=dimension)
        
        st.plotly_chart(fig, use_container_width=True)
        
        pivot = data.pivot_table(index=dimension, columns='Month-Year', values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
        pivot['項目總計'] = pivot.sum(axis=1)
        pivot.loc['當月總計'] = pivot.sum(axis=0)
        st.dataframe(pivot.style.format("{:.0f}"), use_container_width=True)

    # 渲染各維度
    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    # 6. 匯出報告功能
    if st.sidebar.button("📊 導出完整報告"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image:
            slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width)
        buf_ppt = io.BytesIO()
        prs.save(buf_ppt)
        st.sidebar.download_button("下載 PPTX", buf_ppt.getvalue(), "Tactical_Report.pptx")
        
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=16)
        pdf.cell(200, 10, txt="AICS Tactical Deployment Report", ln=True, align='C')
        buf_pdf = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
        st.sidebar.download_button("下載 PDF", buf_pdf, "Tactical_Report.pdf")

else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
