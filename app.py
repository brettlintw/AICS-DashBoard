import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
import datetime

# 1. 初始化頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.2")

st.title("🌐 AICS 北美部署決策中心 (V6.2 客戶維度強化版)")

# 美國 50 州中心座標資料庫 (用於地圖背景標註)
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
    # 載入原始資料
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    df['Month-Year'] = df['Date(出庫)'].dt.strftime('%Y-%m') 

    # 側邊欄過濾控制
    st.sidebar.markdown("---")
    machine_list = df['Machine Type'].unique().tolist()
    selected_machines = st.sidebar.multiselect("設備類型選擇", machine_list, default=machine_list)
    
    date_range = st.sidebar.date_input("日期區間", value=(df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()))
    
    # 執行過濾邏輯
    f_df = df[
        (df['Machine Type'].isin(selected_machines)) & 
        (df['Date(出庫)'].dt.date >= date_range[0]) & 
        (df['Date(出庫)'].dt.date <= date_range[1])
    ].copy()

    # 3. 戰報摘要：總量統計
    st.markdown("---")
    col_total, col_each = st.columns([1, 3])
    with col_total:
        st.subheader("📊 區間總出貨量")
        st.markdown(f"<h1 style='color:#e63946; font-size:65px;'>{int(f_df['Outbound Qty (Item)'].sum())} <small style='font-size:20px;'>台</small></h1>", unsafe_allow_html=True)
    
    with col_each:
        st.subheader("📱 設備即時統計")
        m_cols = st.columns(len(selected_machines))
        for idx, m_name in enumerate(selected_machines):
            m_sum = f_df[f_df['Machine Type'] == m_name]['Outbound Qty (Item)'].sum()
            m_cols[idx].metric(m_name, f"{int(m_sum)} 台")

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

    # 5. 餅圖結構分析 (新增客戶占比)
    st.markdown("---")
    st.subheader("🍕 結構占比分析")
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        st.plotly_chart(px.pie(f_df, values='Outbound Qty (Item)', names='Machine Type', title="設備總占比", hole=0.4), use_container_width=True)
    with p2:
        st.plotly_chart(px.pie(f_df, values='Outbound Qty (Item)', names='Field', title="場域(Field)占比"), use_container_width=True)
    with p3:
        st.plotly_chart(px.pie(f_df, values='Outbound Qty (Item)', names='Area', title="區域(Area)占比"), use_container_width=True)
    with p4:
        st.plotly_chart(px.pie(f_df, values='Outbound Qty (Item)', names='Company', title="客戶(Company)占比"), use_container_width=True)

    # --- 6. 多維度推移分析模組 ---
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name} 每月出貨推移圖與明細矩陣")
        
        if not data.empty:
            # A. 繪製推移圖
            trend_df = data.groupby(['Month-Year', dimension])['Outbound Qty (Item)'].sum().reset_index()
            trend_df = trend_df.sort_values('Month-Year')
            
            fig_line = px.line(trend_df, x='Month-Year', y='Outbound Qty (Item)', color=dimension, 
                               markers=True, text='Outbound Qty (Item)', title=f"{title_name} 推移趨勢")
            fig_line.update_traces(textposition="top center", line=dict(width=4, shape='spline'))
            st.plotly_chart(fig_line, use_container_width=True)
            
            # B. 建立明細矩陣
            pivot = data.pivot_table(
                index=dimension, 
                columns='Month-Year', 
                values='Outbound Qty (Item)', 
                aggfunc='sum', 
                fill_value=0
            )
            pivot['項目總計'] = pivot.sum(axis=1)
            pivot.loc['當月總計'] = pivot.sum(axis=0)
            
            st.markdown(f"#### {title_name} 數據明細表")
            st.dataframe(pivot.style.format("{:.0f}"), use_container_width=True)
        else:
            st.warning(f"當前篩選條件下無 {title_name} 數據。")

    # 執行維度渲染 (新增客戶維度)
    render_analysis_section(f_df, 'Machine Type', '設備維度 (Machine Type)')
    render_analysis_section(f_df, 'Field', '場域維度 (Field)')
    render_analysis_section(f_df, 'Area', '區域維度 (Area)')
    render_analysis_section(f_df, 'Company', '客戶維度 (Company)')

    # 7. PowerPoint 報告匯出
    if st.sidebar.button("📊 生成專業戰報 (PPTX)"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image:
            slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        rows, cols = len(selected_machines) + 2, 2
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(2), Inches(4), Inches(2)).table
        table.cell(0,0).text = "設備/分類"; table.cell(0,1).text = "總出貨量"
        for i, m_name in enumerate(selected_machines):
            total = f_df[f_df['Machine Type'] == m_name]['Outbound Qty (Item)'].sum()
            table.cell(i+1, 0).text = m_name
            table.cell(i+1, 1).text = str(int(total))
        table.cell(rows-1, 0).text = "總計"
        table.cell(rows-1, 1).text = str(int(f_df['Outbound Qty (Item)'].sum()))
        
        buf = io.BytesIO()
        prs.save(buf)
        st.sidebar.download_button("📥 下載戰報檔案", buf.getvalue(), f"AICS_Tactical_Report_{datetime.date.today()}.pptx")

else:
    st.info("💡 指揮中心就緒，請在側邊欄上傳 Excel 數據檔案。")
