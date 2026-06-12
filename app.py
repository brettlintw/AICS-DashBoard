import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V6.7")

# CSS 樣式設定
st.markdown("""
    <style>
    .legend-text { font-size: 28px !important; }
    .blue-cell { color: #0000ff !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🌐 AICS 北美部署決策中心 (V6.7 最終優化版)")

# 數據導入
uploaded_file = st.sidebar.file_uploader("上傳 Excel 數據", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳戰報背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    # 日期區間篩選
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1])]
    
    df['Month-Year'] = df['Date(出庫)'].dt.strftime('%Y-%m')
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    f_df = df[df['Machine Type'].isin(selected_machines)].copy()

    # 1. 設備統計表 (含合計)
    st.subheader("📊 設備總覽統計")
    summary_data = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    summary_data.loc[len(summary_data)] = ['合計', summary_data['Outbound Qty (Item)'].sum()]
    st.table(summary_data)

    # 2. 北美地圖模組 (新增州別代碼顯示 & 圖例放大)
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    
    # 顯示州代碼的邏輯 (使用 Scattergeo 的 text 模式)
    if not f_df.empty:
        # 先畫圓點數據
        for m in selected_machines:
            d = f_df[f_df['Machine Type'] == m].groupby('State Code')['Outbound Qty (Item)'].sum().reset_index()
            fig_map.add_trace(go.Scattergeo(locations=d['State Code'], locationmode="USA-states", 
                                            marker=dict(size=d['Outbound Qty (Item)']*1.5), name=m))
        
        # 額外新增一層文字標註，顯示州代碼
        map_labels = f_df.groupby('State Code').first().reset_index()
        fig_map.add_trace(go.Scattergeo(locations=map_labels['State Code'], locationmode="USA-states", 
                                        text=map_labels['State Code'], mode='text', 
                                        textfont=dict(color='black', size=12), showlegend=False))
        
    fig_map.update_layout(geo=dict(scope='usa'), height=600, margin={"r":0,"t":0,"l":0,"b":0}, 
                          legend=dict(font=dict(size=20))) # 圖例字體放大
    st.plotly_chart(fig_map, use_container_width=True)

    # 3. 分析模組 (含藍色字體、合計列、縮放開關)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        
        mode = st.radio(f"檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"mode_{dimension}")
        chart_type = st.radio(f"圖表類型", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=f"chart_{dimension}")
        
        # 數據計算
        if mode == "月份推移":
            df_g = data.groupby(['Month-Year', dimension])['Outbound Qty (Item)'].sum().reset_index()
            pivot = data.pivot_table(index=dimension, columns='Month-Year', values='Outbound Qty (Item)', aggfunc='sum', fill_value=0, margins=True, margins_name='合計')
        else:
            df_g = data.groupby(dimension)[['Outbound Qty (Item)']].sum().reset_index()
            pivot = df_g.set_index(dimension)
            pivot.loc['合計'] = pivot.sum()

        # 繪圖 (柱狀圖寬度優化)
        if chart_type == "推移圖":
            fig = px.line(df_g, x='Month-Year' if mode=="月份推移" else dimension, y='Outbound Qty (Item)', color=dimension, markers=True, text='Outbound Qty (Item)')
        elif chart_type == "柱狀圖":
            fig = px.bar(df_g, x='Month-Year' if mode=="月份推移" else dimension, y='Outbound Qty (Item)', color=dimension, text='Outbound Qty (Item)')
            fig.update_layout(bargap=0.3)
        else:
            fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
            
        st.plotly_chart(fig, use_container_width=True)

        # 藍色字體表格開關
        if st.checkbox(f"顯示 {title_name} 數據列表", key=f"check_{dimension}"):
            st.dataframe(pivot.style.map(lambda x: 'color: blue; font-weight: bold;'), use_container_width=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    # 4. 導出
    if st.sidebar.button("📊 導出完整報告"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image: slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width)
        buf = io.BytesIO(); prs.save(buf)
        st.sidebar.download_button("下載 PPTX", buf.getvalue(), "Tactical_Report.pptx")
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
