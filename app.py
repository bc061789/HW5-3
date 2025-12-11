import streamlit as st
import sys
import subprocess
from io import BytesIO

# 🔧 確保 python-pptx 有安裝，沒有就現場 pip install
def ensure_pptx():
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
    except ModuleNotFoundError:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "python-pptx"],
            check=True,
        )
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
    return Presentation, Inches, Pt, RGBColor


st.set_page_config(page_title="AI PPT Re-Designer", page_icon="🧠")
st.title("🧠 AI PowerPoint 版型重新設計 Demo")

st.markdown("""
上傳一份 PPTX，選擇一種風格，產生重新設計後的簡報。
""")

uploaded = st.file_uploader("請上傳 PPTX 檔案", type=["pptx"])
style = st.radio("選擇風格", ["科技藍 Tech Style", "極簡白 Minimal Style"])

if uploaded and st.button("🚀 產生新的 PPT"):
    # ⬇️ 在用到 pptx 前，再呼叫我們的 ensure_pptx
    Presentation, Inches, Pt, RGBColor = ensure_pptx()

    old_prs = Presentation(uploaded)
    new_prs = Presentation()

    # 清空預設投影片
    while len(new_prs.slides) > 0:
        r_id = new_prs.slides._sldIdLst[0].rId
        new_prs.part.drop_rel(r_id)
        del new_prs.slides._sldIdLst[0]

    for old_slide in old_prs.slides:
        slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
        bg_fill = slide.background.fill
        bg_fill.solid()

        if style.startswith("科技藍"):
            bg_fill.fore_color.rgb = RGBColor(8, 24, 72)
            font_color = RGBColor(255, 255, 255)
        else:
            bg_fill.fore_color.rgb = RGBColor(255, 255, 255)
            font_color = RGBColor(40, 40, 40)

        top = Inches(1)

        for shape in old_slide.shapes:
            if not shape.has_text_frame:
                continue

            box = slide.shapes.add_textbox(Inches(1), top, Inches(8), Inches(1))
            tf = box.text_frame
            tf.text = shape.text

            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(24)
                    r.font.color.rgb = font_color

            top += Inches(0.8)

        if style.startswith("極簡白"):
            line = slide.shapes.add_shape(
                autoshape_type_id=1,
                left=Inches(0.8),
                top=Inches(0.8),
                width=Inches(0.05),
                height=Inches(6),
            )
            line.fill.solid()
            line.fill.fore_color.rgb = RGBColor(210, 180, 90)
            line.line.fill.background()

    output = BytesIO()
    new_prs.save(output)
    output.seek(0)

    filename = "tech_style_redesign.pptx" if style.startswith("科技藍") else "minimal_style_redesign.pptx"

    st.success("✅ 重新設計完成！請下載新的 PPT。")
    st.download_button(
        label="💾 下載新 PPT",
        data=output,
        file_name=filename,
        mime="applic
