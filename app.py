import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
import io

# 1. 頁面設定
st.set_page_config(layout="wide", page_title="AICS 北美部署決策中心 V8.11")
st.markdown("""<style>.stDataFrame table td, .stDataFrame table th { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)

st.title("🌐 AICS 北美部署決策中心 (V8.11 排序修復版)")

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

    def render_analysis_section(data, dimension, title_name):
        st.markdown("---")
        st.subheader(f"📈 {title_name}")
        col1, col2 = st.columns([1, 1])
        with col1:
            mode = st.radio("檢視模式", ["月份推移", "全時間段彙總"], horizontal=True, key=f"m_{dimension}")
            chart_type = st.radio("圖表類型", ["推移圖", "柱狀圖", "餅圖"], horizontal=True, key=f"c_{dimension}")
        with col2:
            sort = st.selectbox("排序方式", ["預設", "由大至小", "由小至大"], key=f"s_{dimension}") if mode == "全時間段彙總" else "預設"

        df_g = data.groupby(['Month-Date', dimension] if mode=="月份推移" else dimension)[['Outbound Qty (Item)']].sum().reset_index()
        
        # 圖表繪製
        x_col = 'Month-Date' if mode == "月份推移" else dimension
        if chart_type == "推移圖": fig = px.line(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension, markers=True)
        elif chart_type == "柱狀圖": fig = px.bar(df_g, x=x_col, y='Outbound Qty (Item)', color=dimension); fig.update_layout(bargap=0.3)
        else: fig = px.pie(df_g, values='Outbound Qty (Item)', names=dimension)
        st.plotly_chart(fig, use_container_width=True)

        # 數據列表處理
        if st.checkbox(f"顯示數據列表", key=f"ch_{dimension}"):
            pivot = data.pivot_table(index=dimension, columns='Month-Date' if mode=="月份推移" else None, values='Outbound Qty (Item)', aggfunc='sum', fill_value=0)
            pivot['合計'] = pivot.sum(axis=1)
            
            # 【關鍵修復】：在排序前先分離合計行，排序主數據後再拼回
            main_df = pivot.drop('合計', errors='ignore')
            if mode == "全時間段彙總" and sort != "預設":
                # 對主體數據進行排序，基於第一列或合計列
                sort_col = '合計' if '合計' in main_df.columns else main_df.columns[0]
                main_df = main_df.sort_values(by=sort_col, ascending=(sort=="由小至大"))
            
            # 重新拼回合計
            final_df = pd.concat([main_df, pivot.loc[['合計']] if '合計' in pivot.index else pd.DataFrame()])
            
            # 增加列總計
            col_totals = final_df.sum(axis=0)
            col_totals.name = '總計'
            final_df = pd.concat([final_df, col_totals.to_frame().T])
            
            st.markdown(final_df.style.map(lambda x: 'color: blue; font-weight: bold;' if '總計' in str(x) else 'color: black').to_html(), unsafe_allow_html=True)

    for dim, name in [('Machine Type', '設備維度'), ('Field', '場域維度'), ('Area', '區域維度'), ('Company', '客戶維度'), ('Device/Platform', '平台維度')]:
        render_analysis_section(f_df, dim, name)
else:
    st.info("💡 請上傳數據檔案以啟動戰情室。")
