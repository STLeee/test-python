import matplotlib.pyplot as plt
import io


def latex_to_image(latex_str):
    fig = plt.figure()
    text = fig.text(0, 0, f"${latex_str}$", fontsize=8)
    fig.canvas.draw()
    bbox = text.get_window_extent()
    width, height = bbox.size / fig.dpi + 0.05
    fig.set_size_inches((width, height))
    buffer = io.BytesIO()
    plt.axis('off')
    plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight', pad_inches=0.1)
    plt.close(fig)
    buffer.seek(0)
    return buffer


from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 註冊字體
pdfmetrics.registerFont(TTFont('Microsoft JhengHei Bold', 'microsoft-jhenghei-bold.ttf'))
pdfmetrics.registerFont(TTFont('Microsoft JhengHei', 'microsoft-jhenghei.ttf'))


from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm


def generate_pdf(quizzes, output_path="output.pdf"):
    doc = SimpleDocTemplate(
        output_path, pagesize=A4, rightMargin=20 * mm, leftMargin=20 * mm, topMargin=20 * mm, bottomMargin=20 * mm
    )
    styles = getSampleStyleSheet()
    styles['Title'].fontName = 'Microsoft JhengHei Bold'
    styles['Heading2'].fontName = 'Microsoft JhengHei Bold'
    styles['Normal'].fontName = 'Microsoft JhengHei'
    story = []

    # 標題
    story.append(Paragraph(f"{quizzes['academic_year']} 學年度 {quizzes['semester']} 學期", styles['Title']))
    story.append(
        Paragraph(f"{quizzes['grade']}年級 {quizzes['subject']}科 第{quizzes['chapter']}課 {quizzes['title']}", styles['Title'])
    )
    story.append(Spacer(1, 12))

    # 選擇題
    story.append(Paragraph("壹、選擇題 (每題 ___ 分。共 ____ 分)：", styles['Heading2']))
    for idx, quiz in enumerate([q for q in quizzes['quizzes'] if q['quiz_type'] == 'mcq']):
        story.append(Paragraph(f"{idx + 1}. {quiz['question']}", styles['Normal']))
        for opt_idx, option in enumerate(quiz['options']):
            story.append(Paragraph(f"({chr(65 + opt_idx)}) {option}", styles['Normal']))
        # 如果有 LaTeX 方程式
        if '$' in quiz['question']:
            latex_str = quiz['question'].split('$')[1]
            img_buffer = latex_to_image(latex_str)
            img = Image(img_buffer)
            img._restrictSize(100 * mm, 20 * mm)
            story.append(img)
        story.append(Spacer(1, 12))

    # 簡答題
    story.append(Paragraph("貳、簡答題 (每題 ___ 分。共 ____ 分)：", styles['Heading2']))
    for idx, quiz in enumerate([q for q in quizzes['quizzes'] if q['quiz_type'] == 'saq']):
        story.append(Paragraph(f"{idx + 1}. {quiz['question']}", styles['Normal']))
        story.append(Spacer(1, 24))
        # 如果有 LaTeX 方程式
        if '$' in quiz['question']:
            latex_str = quiz['question'].split('$')[1]
            img_buffer = latex_to_image(latex_str)
            img = Image(img_buffer)
            img._restrictSize(100 * mm, 20 * mm)
            story.append(img)
        story.append(Spacer(1, 12))

    doc.build(story)


import json


def main():
    with open("quizzes.json", "r", encoding="utf-8") as f:
        quizzes = json.load(f)
    generate_pdf(quizzes, output_path="demo.pdf")


if __name__ == "__main__":
    main()
