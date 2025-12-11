import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO

st.set_page_config(page_title="AI PPT Re-Designer", page_icon="🧠")
st.title("🧠 AI PowerPoint 版型重新設計 Demo")

uploaded = st.file_uploader("請上傳 PPTX 檔案", type=["pptx"])
style = st.radio("選擇風格", ["科技藍 Tech Style", "極簡白 Minimal Style"])

if uploaded and st.button("🚀 產生新的 PPT"):
    old = Presentation(uploaded)
    new = Presentation()

    # 移除新簡報預設頁面
    while len(new.slides) > 0:
        rId = new.slides._sldIdLst[0].rId
        new.part.drop_rel(rId)
        del new.slides._sldIdLst[0]

    for old_slide in old.slides:
        slide = new.slides.add_slide(new.slide_layouts[6])
        bg = slide.background.fill
        bg.solid()

        if style.startswith("科技藍"):
            bg.fore_color.rgb = RGBColor(10, 30, 80)
            font_color = RGBColor(255, 255, 255)
        else:
            bg.fore_color.rgb = RGBColor(255, 255, 255)
            font_color = RGBColor(50, 50, 50)

        y = Inches(1)

        for shape in old_slide.shapes:
            if not shape.has_text_frame:
                continue

            box = slide.shapes.add_textbox(Inches(1), y, Inches(8), Inches(1))
            tf = box.text_frame
            tf.text = shape.text

            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(24)
                    r.font.color.rgb = font_color

            y += Inches(0.8)

        if style.startswith("極簡白"):
            line = slide.shapes.add_shape(
                1, Inches(0.8), Inches(0.8), Inches(0.05), Inches(6)
            )
            line.fill.solid()
            line.fill.fore_color.rgb = RGBColor(200, 170, 90)
            line.line.fill.background()

    buf = BytesIO()
    new.save(buf)
    buf.seek(0)

    filename = "tech_style_redesign.pptx" if style.startswith("科技藍") else "minimal_style_redesign.pptx"

    st.success("🎉 已完成重新設計，請下載！")
    st.download_button(
        label="💾 下載新 PPT",
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

elif not uploaded:
    st.info("請先上傳 PPTX 檔案才能開始。")
