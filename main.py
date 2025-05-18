import streamlit as st

st.set_page_config(page_title="出荷分析统一系统", layout="wide")
st.title("📦 出荷統合 × 分析 系统")

st.markdown("""
欢迎使用本系统。

- 左侧菜单可以切换页面：
  - `📂 数据清洗`：上传原始 Excel，进行清洗整合；
  - `📊 数据分析`：对清洗后的结果进行达成率与断货分析。
""")
