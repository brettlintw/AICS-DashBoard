import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io

# 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.3")
st.markdown("""<style>.stDataFrame table td, .stDataFrame table th { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)

st.title("🌐 AICS 北美部署決策中心 (V8.3 數據同步與補全版)")

# 數據導入
st.sidebar.header("⚙️ 戰情控制台")
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    
    # 確保 f_df 包含篩選後的所有數據
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # 1. 設備總覽統計 (強制聯動)
    st.subheader("📊 設備總覽統計")
    summary = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    total_row = pd.DataFrame({'Machine Type': ['合計'], 'Outbound Qty (Item)': [summary['Outbound Qty (Item)'].sum()]})
    summary = pd.concat([total_row, summary], ignore_index=True)
    st.markdown(summary.style.map(lambda x: 'color: blue; font-weight: bold;' if x == '合計' else 'color: black').to_html(), unsafe_allow_html=True)

    # 2. 北美地圖 (地區補全邏輯)
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    all_states = sorted(df['State Code'].unique()) # 取得所有可能的地區
    
    for m in selected_machines:
        # 強制補全：先建立完整地區表，再 merge 數據，沒數據的補 0
        m_df = f_df[f_df['Machine Type'] == m].groupby('State Code')['Outbound Qty (Item)'].sum().reindex(all_states, fill_value=0).reset_index()
        fig_map.add_trace(go.Scattergeo(locations=m_df['State Code'], locationmode="USA-states", marker=dict(size=m_df['Outbound Qty (Item)']*1.5), name=m))
    
    fig_map.update_layout(geo=dict(scope='north america', projection_type='albers usa', center=dict(lat=38, lon=-100), projection_scale=1.5), height=800, margin={"r":0,"t":0,"l":0,"b":0}, legend=dict(font=dict(size=20)))
    st.plotly_chart(fig_map, use_container_width=True)

    # 3. 分析模組 (數據列表強綁定)
    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        mode = st.radio("檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"m_{dimension}")
        chart_type = st.radio("圖表類型", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=f"c_{dimension}")
        sort = st.selectbox("排序方式", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}") if mode == "全時間段彙總" else "預設"

        df_g = data.groupby(['Month-Date', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
        if mode == "全時間段彙總" and sort != "預設":
            df_g = df_g.sort_values(by='Outbound Qty (Item)', ascending=(sort=="由小至大"))

        x_col = 'Month-Date' if mode == "月份推移" else dimension
        if chart_type == "推移圖": fig = px.line(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension, markers=True, text='Outbound Qty (Item)')
        elif chart_type == "柱狀圖": fig = px.bar(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension, text='Outbound Qty (Item)'); fig.update_layout(bargap=0.3)
        else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
        st.plotly_chart(fig, use_container_width=True)

        # 這裡直接使用當前篩選後的 data
        if st.checkbox(f"顯示數據列表", key=f"ch_{dimension}"):
            pivot = data.pivot_table(index=dimension, columns='Month-Date' if mode=="月份推移" else None, values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
            total = pivot.sum(axis=0); total.name = '合計'
            pivot = pd.concat([total.to_frame().T, pivot])
            st.markdown(pivot.style.map(lambda x: 'color: blue; font-weight: bold;' if '合計' in str(x) else 'color: black').to_html(), unsafe_allow_html=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)

    if st.sidebar.button("📊 導出 PPTX"):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if bg_image: slide.shapes.add_picture(io.BytesIO(bg_image.read()), 0, 0, width=prs.slide_width)
        buf = io.BytesIO(); prs.save(buf)
        st.sidebar.download_button("下載 PPTX", buf.getvalue(), "Tactical_Report.pptx")
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
