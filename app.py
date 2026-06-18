import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
import plotly.io as pio
import streamlit.components.v1 as components

# 頁面設定：強制列印佈局與色彩優化
st.set_page_config(layout="wide", page_title="AICS 北美決策中心")

# [終極列印 CSS] 強制色彩飽和度與不透明度，徹底解決霧白
st.markdown("""
    <style>
    @media print {
        /* 1. 強制隱藏所有 UI 控制元件 */
        [data-testid="stSidebar"], .stButton, .stFileUploader, .stRadio, .stSelectbox, .stCheckbox { 
            display: none !important; 
        }
        
        /* 2. 強制文字變更為純黑，確保列印清晰度 */
        body, p, h1, h2, h3, .stDataFrame, .stMarkdown { 
            color: #000000 !important; 
            -webkit-text-fill-color: #000000 !important; 
        }
        
        /* 3. 強制圖表渲染：取消淡化，拉高對比與飽和度以抵銷列印時的色階流失 */
        .stPlotlyChart { 
            opacity: 1.0 !important; 
            filter: contrast(180%) brightness(100%) saturate(150%) !important; 
            page-break-inside: avoid !important; 
        }
        
        /* 4. 強制瀏覽器使用原始顏色空間，不執行灰階轉換 */
        body { 
            -webkit-print-color-adjust: exact !important; 
            print-color-adjust: exact !important; 
            background-color: #ffffff !important;
        }
        
        /* 5. 撐開滿版頁面 */
        [data-testid="block-container"] { max-width: 100% !important; padding: 10px !important; }
    }
    </style>
""", unsafe_allow_html=True)

st.title("🌐 AICS 北美部署決策中心 (V8.28 最終色彩校正版)")

# 核心顯示函數：強制 Plotly 圖表背景為不透明白色
def display_chart(fig):
    fig.update_layout(
        paper_bgcolor='rgba(255,255,255,1)', 
        plot_bgcolor='rgba(255,255,255,1)',
        margin=dict(l=20, r=20, t=30, b=20)
    )
    st.plotly_chart(fig, use_container_width=True)

# 北美州代碼映射
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    st.sidebar.markdown("### 匯出報告")
    if st.sidebar.button("🖨️ 匯出 PDF 報告"):
        components.html("<script>window.parent.window.print();</script>", height=0)
    
    date_range = st.sidebar.date_input("日期篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    
    filters = {}
    dims = [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]
    for col, label in dims:
        selected = st.sidebar.multiselect(f"{label}篩選 (留空即全選)", df[col].unique().tolist())
        filters[col] = selected if selected else df[col].unique().tolist()
    
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1])].copy()
    for col, selected in filters.items():
        f_df = f_df[f_df[col].isin(selected)]
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # 地圖渲染
    st.subheader("🗺️ 北美設備戰術分佈")
    fig_map = go.Figure()
    fig_map.add_trace(go.Scattergeo(lon=[US_STATES_COORDS[s][1] for s in US_STATES_COORDS], lat=[US_STATES_COORDS[s][0] for s in US_STATES_COORDS], text=list(US_STATES_COORDS.keys()), mode='text', textfont=dict(size=14, color='blue'), showlegend=False))
    for m in f_df['Machine Type'].unique():
        m_df = f_df[f_df['Machine Type'] == m].groupby('State Code')['Outbound Qty (Item)'].sum().reset_index()
        fig_map.add_trace(go.Scattergeo(locations=m_df['State Code'], locationmode="USA-states", marker=dict(size=m_df['Outbound Qty (Item)']*2.5), name=m))
    fig_map.update_layout(geo=dict(scope='usa', fitbounds="locations"), height=600, margin={"l": 0, "r": 0, "t": 0, "b": 0})
    display_chart(fig_map)

    # 維度分析
    for dim, name in dims:
        st.markdown("---")
        st.subheader(f"📈 {name}")
        c1, c2 = st.columns([1, 4])
        with c1:
            mode = st.radio("模式", ["月份推移", "全時間段"], key=f"m_{dim}")
            chart_type = st.radio("圖表", ["柱狀圖", "推移圖", "餅圖"], key=f"c_{dim}")
        with c2:
            df_g = f_df.groupby(['Month-Date', dim] if mode=="月份推移" else dim)[['Outbound Qty (Item)']].sum().reset_index()
            if chart_type == "柱狀圖": fig = px.bar(df_g, x=dim if mode=="全時間段" else "Month-Date", y='Outbound Qty (Item)', color=dim)
            elif chart_type == "推移圖": fig = px.line(df_g, x="Month-Date", y='Outbound Qty (Item)', color=dim, markers=True)
            else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dim)
            display_chart(fig)
            if st.checkbox("顯示表格", key=f"t_{dim}"): st.dataframe(df_g, use_container_width=True)

    # PPTX 匯出
    if st.sidebar.button("📊 導出完整戰情室報表"):
        prs = Presentation()
        for dim, name in dims:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.text = f"分析維度: {name}"
            df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
            try:
                fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)')
                img_bytes = pio.to_image(fig, format="png", width=1200, height=450)
                slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(1), width=Inches(9))
            except: pass
            table = slide.shapes.add_table(len(df_g)+1, len(df_g.columns), Inches(0.5), Inches(5), Inches(9), Inches(2)).table
            for i, col in enumerate(df_g.columns): table.cell(0, i).text = str(col)
            for r in range(len(df_g)):
                for c in range(len(df_g.columns)): table.cell(r+1, c).text = str(df_g.iloc[r, c])
        buf = io.BytesIO(); prs.save(buf); st.sidebar.download_button("確認下載 PPTX", buf.getvalue(), "Tactical_Report.pptx")
