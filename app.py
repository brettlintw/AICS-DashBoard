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

st.title("🌐 AICS 北美部署決策中心 (V8.10 專業報表版)")

# 座標映射表
US_STATES_COORDS = {'AL': [32.8, -86.7], 'AK': [61.3, -152.4], 'AZ': [33.7, -111.4], 'AR': [34.9, -92.3], 'CA': [36.1, -119.6], 'CO': [39.0, -105.3], 'CT': [41.5, -72.7], 'DE': [39.3, -75.5], 'FL': [27.7, -81.6], 'GA': [33.0, -83.6], 'HI': [21.0, -157.4], 'ID': [44.2, -114.4], 'IL': [40.3, -88.9], 'IN': [39.8, -86.2], 'IA': [42.0, -93.2], 'KS': [38.5, -96.7], 'KY': [37.6, -84.6], 'LA': [31.1, -91.8], 'ME': [44.6, -69.3], 'MD': [39.0, -76.8], 'MA': [42.2, -71.5], 'MI': [43.3, -84.5], 'MN': [45.6, -93.9], 'MS': [32.7, -89.6], 'MO': [38.4, -92.2], 'MT': [46.9, -110.4], 'NE': [41.1, -98.2], 'NV': [38.3, -117.0], 'NH': [43.4, -71.5], 'NJ': [40.2, -74.5], 'NM': [34.8, -106.2], 'NY': [42.1, -74.9], 'NC': [35.6, -79.8], 'ND': [47.5, -99.7], 'OH': [40.3, -82.7], 'OK': [35.5, -96.9], 'OR': [44.5, -122.0], 'PA': [40.5, -77.2], 'RI': [41.6, -71.5], 'SC': [33.8, -80.9], 'SD': [44.2, -99.4], 'TN': [35.7, -86.6], 'TX': [31.0, -97.5], 'UT': [40.1, -111.8], 'VT': [44.0, -72.7], 'VA': [37.7, -78.1], 'WA': [47.4, -120.4], 'WV': [38.4, -80.9], 'WI': [44.2, -89.6], 'WY': [42.7, -107.3]}

# 數據導入
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])
bg_image = st.sidebar.file_uploader("上傳背景圖", type=["png", "jpg"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    st.subheader("📊 設備總覽統計")
    summary = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    total_row = pd.DataFrame({'Machine Type': ['合計'], 'Outbound Qty (Item)': [summary['Outbound Qty (Item)'].sum()]})
    summary = pd.concat([total_row, summary], ignore_index=True)
    st.dataframe(summary.style.apply(lambda row: ['color: blue; font-weight: bold;' if row['Machine Type'] == '合計' else ''] * len(row), axis=1), use_container_width=True)

    # 專業版 PPTX 導出邏輯
    if st.sidebar.button("📊 導出完整戰情室報表"):
        try:
            prs = Presentation()
            # 定義分析維度
            dimensions = [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]
            
            for dim, name in dimensions:
                slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白頁面
                
                # 1. 寫入標題
                title = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
                title.text_frame.text = f"分析維度: {name}"
                title.text_frame.paragraphs[0].font.bold = True
                title.text_frame.paragraphs[0].font.size = Pt(32)
                
                # 2. 生成並寫入柱狀圖
                df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
                fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)', title=f"{name} 統計")
                
                # 處理圖表轉換 (若環境無 Kaleido，系統會拋出錯誤由下方的 except 捕捉)
                img_bytes = pio.to_image(fig, format="png", width=800, height=350)
                slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(1.2), width=Inches(9))
                
                # 3. 寫入數據表格
                rows, cols = df_g.shape
                table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(5), Inches(9), Inches(2.5)).table
                for i, col_name in enumerate(df_g.columns): table.cell(0, i).text = str(col_name)
                for r in range(rows):
                    for c in range(cols): table.cell(r + 1, c).text = str(df_g.iloc[r, c])
            
            # 背景圖處理
            if bg_image:
                bg_image.seek(0)
                for s in prs.slides: s.shapes.add_picture(bg_image, 0, 0, width=prs.slide_width, height=prs.slide_height).z_order = -1
            
            buf = io.BytesIO()
            prs.save(buf)
            st.sidebar.download_button("下載正式報表 PPTX", buf.getvalue(), "Tactical_Report_Visual.pptx")
            st.sidebar.success("導出成功！")
        except Exception as e:
            st.sidebar.error(f"導出錯誤，請確認 Kaleido 已安裝: {e}")

else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
