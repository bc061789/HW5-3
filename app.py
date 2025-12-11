from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# 共用：原始簡報結構內容
slides_content = [
    ("AI 在零售業的應用", "從需求預測到智慧販賣機"),
    ("AI 應用概述", "．需求預測\n．推薦系統\n．自動補貨"),
    ("需求預測流程", "1. 資料蒐集\n2. 資料清洗\n3. 模型訓練\n4. 上線部署"),
    ("案例：智慧販賣機", "依據天氣、時間與銷售紀錄，自動推薦與補貨"),
    ("結語", "AI 正在改變零售業的營運模式與顧客體驗")
]

# 版本 A：科技藍風格
prs_tech = Presentation()
for title_text, body_text in slides_content:
    slide_layout = prs_tech.slide_layouts[6]  # blank
    slide = prs_tech.slides.add_slide(slide_layout)

    # 深藍背景
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(8, 24, 72)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(8.5), Inches(1))
    tf_title = title_box.text_frame
    p = tf_title.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    # 內文
    body_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(8.2), Inches(4))
    tf_body = body_box.text_frame
    tf_body.word_wrap = True
    p_body = tf_body.paragraphs[0]
    run_body = p_body.add_run()
    run_body.text = body_text
    run_body.font.size = Pt(24)
    run_body.font.color.rgb = RGBColor(220, 230, 255)

prs_tech.save("/mnt/data/retail_ai_tech_style.pptx")

# 版本 B：極簡白風格
prs_min = Presentation()
for title_text, body_text in slides_content:
    slide_layout = prs_min.slide_layouts[6]  # blank
    slide = prs_min.slides.add_slide(slide_layout)

    # 白底
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)

    # 左側細線裝飾
    line = slide.shapes.add_shape(
        1,  # rectangle
        Inches(0.7),
        Inches(0.7),
        Inches(0.05),
        Inches(5.5)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(210, 180, 90)
    line.line.fill.background()

    # 標題
    title_box = slide.shapes.add_textbox(Inches(1.1), Inches(0.7), Inches(8.0), Inches(1))
    tf_title = title_box.text_frame
    p = tf_title.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = RGBColor(40, 40, 40)

    # 內文
    body_box = slide.shapes.add_textbox(Inches(1.3), Inches(1.7), Inches(7.5), Inches(4.5))
    tf_body = body_box.text_frame
    tf_body.word_wrap = True
    p_body = tf_body.paragraphs[0]
    run_body = p_body.add_run()
    run_body.text = body_text
    run_body.font.size = Pt(22)
    run_body.font.color.rgb = RGBColor(70, 70, 70)

prs_min.save("/mnt/data/retail_ai_minimal_style.pptx")

"/mnt/data/retail_ai_tech_style.pptx", "/mnt/data/retail_ai_minimal_style.pptx"
