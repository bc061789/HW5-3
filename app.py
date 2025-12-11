import streamlit as st

st.set_page_config(page_title="AI PPT Re-Designer", page_icon="🧠")
st.title("🧠 AI PowerPoint 版型重新設計 Demo")

uploaded = st.file_uploader("請上傳您的原始 PPTX 檔案", type=["pptx"])
style = st.radio("請選擇 AI 要重新設計的風格", ["科技藍 Tech Theme", "極簡白 Minimal Theme"])

if uploaded:
    st.success(f"已上傳：{uploaded.name}")

    if st.button("✨ 產生新版 AI PPT"):
        if style == "科技藍 Tech Theme":
            path = "ppt/retail_ai_tech_style.pptx"
            filename = "AI_redesign_tech_style.pptx"
            label = "💾 下載科技藍風格新 PPT"
        else:
            path = "ppt/retail_ai_minimal_style.pptx"
            filename = "AI_redesign_minimal_style.pptx"
            label = "💾 下載極簡白風格新 PPT"

        with open(path, "rb") as f:
            data = f.read()

        st.success("🎉 AI 已完成重新設計！")
        st.download_button(
            label=label,
            data=data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

else:
    st.info("請先上傳原始 PPTX。")
