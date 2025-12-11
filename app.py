import streamlit as st
from io import BytesIO

st.set_page_config(page_title="AI PPT Re-Designer", page_icon="🧠")
st.title("🧠 AI PowerPoint 版型重新設計 Demo")

st.markdown("""
這個 demo 示範：

1. 上傳一份原始 PPTX
2. 選擇一種 AI 版型風格
3. 由 AI 給出排版建議與說明（不直接修改檔案）

實際重新設計後的兩種 PPT 檔案，已在報告中另外提供。
""")

uploaded = st.file_uploader("請上傳 PPTX 檔案", type=["pptx"])

style = st.radio("選擇想要的風格", ["科技藍 Tech Style", "極簡白 Minimal Style"])

if uploaded:
    st.success("✅ 已上傳檔案：{}".format(uploaded.name))

    if st.button("✨ 產生 AI 排版建議"):
        # 讀檔大小只是做個小展示，證明有真的收到檔案
        file_bytes = uploaded.getbuffer()
        st.write(f"檔案大小：約 {len(file_bytes) / 1024:.1f} KB")

        if style == "科技藍 Tech Style":
            st.subheader("🎨 科技藍 Tech Style 排版建議")
            st.markdown("""
- **配色**：深藍＋白色文字，加入一點漸層或霓虹感線條做科技感背景  
- **標題頁**：大標題置中，底部加上細線或微光效果  
- **內容頁**：每一頁最多 3 個重點，搭配簡單 icon  
- **流程頁**：用水平流程圖（Step1~4），每個步驟用圓角方塊＋淡光暈  
- **結語頁**：保留大量留白，只放一句總結句搭配小圖示
            """)
        else:
            st.subheader("🧼 極簡白 Minimal Style 排版建議")
            st.markdown("""
- **配色**：純白背景＋深灰文字，點綴一點米色或淺金色線條  
- **標題頁**：左上對齊標題，右下角一條細線做裝飾  
- **內容頁**：文字靠左、圖示小小的放在文字前，不使用粗重的框線  
- **流程頁**：垂直排列 4 個步驟，使用編號 1–4 + 簡短描述  
- **結語頁**：一行簡短總結文字＋很小的 logo 或圖示，整頁幾乎都是留白
            """)

        st.info("""
👉 實作說明：  
此 Demo 由 Streamlit + ChatGPT 產生版型建議文字。  
實際重新設計後的兩種 PPT 檔案（科技藍版、極簡白版），
是依照這些建議在 PowerPoint 中完成，並附在報告與 GitHub Repo。
""")
else:
    st.info("請先上傳一份 PPTX 檔案。")
