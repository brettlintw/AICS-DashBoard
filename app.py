import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import plotly.io as pio

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.10")
st.markdown("""<style>.stDataFrame table td, .stDataFrame table th { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)

st.title("🌐 AICS 北美部署決策中心 (V8.10 穩定版)")

# 座標映射表
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

# 數據導入
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # 1. 設備總覽
    st.subheader("📊 設備總覽統計")
    summary = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    st.dataframe(summary, use_container_width=True)

    # 2. 地圖渲染 (補回地圖與區域邏輯)
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    for state, coords in US_STATES_COORDS.items():
        fig_map.add_trace(go.Scattergeo(lon=[coords[1]], lat=[coords[0]], text=[state], mode='text', textfont=dict(size=9, color='darkblue'), hoverinfo='skip', showlegend=False))
    for m in selected_machines:
        m_df = f_df[f_df['Machine Type'] == m].groupby('State Code')['Outbound Qty (Item)'].sum().reset_index()
        fig_map.add_trace(go.Scattergeo(locations=m_df['State Code'], locationmode="USA-states", marker=dict(size=m_df['Outbound Qty (Item)']*1.5, opacity=0.5), name=m))
    fig_map.update_layout(geo=dict(scope='north america', projection_type='albers usa', center=dict(lat=38, lon=-100), projection_scale=1.5), height=600, margin={"r":0,"t":0,"l":0,"b":0})
    st.plotly_chart(fig_map, use_container_width=True)

    # 3. 分析模組 (補回分析維度並同步 UI)
    def render_analysis_section(dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        df_g = f_df.groupby(dimension)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = px.bar(df_g, x=dimension, y='Outbound Qty (Item)', text='Outbound Qty (Item)')
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(df_g, use_container_width=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(dim, name)

    # 4. PPTX 導出 (維持專業報表佈局)
    if st.sidebar.button("📊 導出完整戰情室報表"):
        try:
            prs = Presentation()
            dims = [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]
            for dim, name in dims:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                title = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
                title.text_frame.text = f"分析維度: {name}"
                
                df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
                
                try:
                    fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)')
                    img_bytes = pio.to_image(fig, format="png")
                    slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(1), width=Inches(9))
                except Exception as e:
                    slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(8), Inches(1)).text_frame.text = f"圖表渲染提示: {str(e)}"
                
                rows, cols = df_g.shape
                table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(4.5), Inches(9), Inches(2)).table
                for i, col in enumerate(df_g.columns): table.cell(0, i).text = str(col)
                for r in range(rows):
                    for c in range(cols): table.cell(r + 1, c).text = str(df_g.iloc[r, c])
            
            buf = io.BytesIO()
            prs.save(buf)
            st.sidebar.download_button("下載報表", buf.getvalue(), "Tactical_Report.pptx")
            st.sidebar.success("導出完成！")
        except Exception as e:
            st.sidebar.error(f"導出錯誤: {e}")
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
