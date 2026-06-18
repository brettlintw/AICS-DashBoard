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
st.title("🌐 AICS 北美部署決策中心 (專業報表版)")

# 數據導入模組
uploaded_file = st.sidebar.file_uploader("上傳 Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Data Base')
    df.columns = df.columns.str.strip()
    df['Date(出庫)'] = pd.to_datetime(df['Date(出庫)'])
    
    date_range = st.sidebar.date_input("日期區間篩選", [df['Date(出庫)'].min().date(), df['Date(出庫)'].max().date()])
    selected_machines = st.sidebar.multiselect("設備類型篩選", df['Machine Type'].unique(), default=df['Machine Type'].unique())
    
    f_df = df[(df['Date(出庫)'].dt.date >= date_range[0]) & (df['Date(出庫)'].dt.date <= date_range[1]) & (df['Machine Type'].isin(selected_machines))].copy()
    f_df['Month-Date'] = f_df['Date(出庫)'].dt.strftime('%Y-%m')

    # UI 渲染區 (設備總覽與分析模組)
    st.subheader("📊 設備總覽統計")
    summary = f_df.groupby('Machine Type')['Outbound Qty (Item)'].sum().reset_index()
    st.dataframe(summary, use_container_width=True)

    # 4. 專業版 PPTX 導出邏輯 (上半部圖表 + 下半部表格)
    if st.sidebar.button("📊 導出完整戰情室報表"):
        try:
            prs = Presentation()
            dims = [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]
            
            for dim, name in dims:
                # 建立一張空白投影片
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                # 1. 寫入標題 (設定於頂部)
                title = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(9), Inches(0.5))
                title.text_frame.text = f"分析維度: {name}"
                title.text_frame.paragraphs[0].font.size = Pt(28)
                
                # 2. 生成統計數據 (用於圖表與表格)
                df_g = f_df.groupby(dim)[['Outbound Qty (Item)']].sum().reset_index().sort_values(by='Outbound Qty (Item)', ascending=False)
                
                # 3. 繪製圖表並插入上半部
                fig = px.bar(df_g, x=dim, y='Outbound Qty (Item)', title=f"{name} 趨勢")
                try:
                    # 快照渲染
                    img_bytes = pio.to_image(fig, format="png", width=1200, height=500)
                    slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(0.7), width=Inches(9), height=Inches(3.5))
                except Exception as e:
                    st.sidebar.warning(f"圖表渲染失敗: {e}")

                # 4. 插入下半部表格
                rows, cols = df_g.shape
                # 定位在 Inches(4.5) 以下，避免與圖表重疊
                table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(4.5), Inches(9), Inches(2)).table
                for i, col_name in enumerate(df_g.columns): table.cell(0, i).text = str(col_name)
                for r in range(rows):
                    for c in range(cols): table.cell(r + 1, c).text = str(df_g.iloc[r, c])
            
            # 下載邏輯
            buf = io.BytesIO()
            prs.save(buf)
            st.sidebar.download_button("下載正式報表 PPTX", buf.getvalue(), "Tactical_Report_Pro.pptx")
            st.sidebar.success("報表生成完畢！")
        except Exception as e:
            st.sidebar.error(f"導出異常: {e}")
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
