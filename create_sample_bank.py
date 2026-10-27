import os
from docx import Document
from docx.shared import RGBColor


def create_question_bank():
  doc = Document()
  doc.add_heading('NGÂN HÀNG CÂU HỎI THI THỬ', level=0)

  # HEADING 1: I. TRẮC NGHIỆM
  doc.add_heading('I. Trắc Nghiệm', level=1)
  for i in range(1, 101):  # Tạo mô phỏng 100 câu trắc nghiệm
    doc.add_paragraph(
        f'Câu {i}: Khái niệm cơ bản số {i} trong chương trình là gì?'
    )
    doc.add_paragraph('A. Phương án A')

    p_correct = doc.add_paragraph()
    run = p_correct.add_run('B. Đáp án đúng (được tô đỏ)')
    run.font.color.rgb = RGBColor(255, 0, 0)

    doc.add_paragraph('*C. Đáp án đúng (được đánh dấu *)')
    doc.add_paragraph('D. Phương án D')
    doc.add_paragraph('')

  # HEADING 1: II. ĐÚNG SAI
  doc.add_heading('II. Đúng Sai', level=1)
  for i in range(1, 30):  # Tạo mô phỏng 30 câu đúng sai
    doc.add_paragraph(f'Câu {i}: Cho các phát biểu sau về bài học số {i}:')

    p_a = doc.add_paragraph()
    r_a = p_a.add_run('a) Phát biểu 1 đúng')
    r_a.font.color.rgb = RGBColor(255, 0, 0)

    doc.add_paragraph('b) Phát biểu 2 sai')
    doc.add_paragraph('*c) Phát biểu 3 đúng')
    doc.add_paragraph('d) Phát biểu 4 sai')
    doc.add_paragraph('')

  # HEADING 1: III. TỰ LUẬN
  doc.add_heading('III. Tự luận', level=1)
  for i in range(1, 20):  # Tạo mô phỏng 20 câu tự luận
    doc.add_paragraph(
        f'Câu {i}: Phân tích tầm quan trọng của chủ đề {i} trong thực tiễn.'
    )
    doc.add_paragraph(
        f'Đáp án mẫu: Chủ đề {i} đóng vai trò quan trọng trong việc tối ưu hóa'
        ' quy trình, nâng cao hiệu suất làm việc và phát triển tư duy hệ thống.'
    )
    doc.add_paragraph('')

  os.makedirs('ngan_hang_de', exist_ok=True)
  output_path = os.path.join('ngan_hang_de', 'Ngan_Hang_Cau_Hoi.docx')
  doc.save(output_path)
  print(f'Đã khởi tạo xong Ngân hàng câu hỏi mẫu tại: {output_path}')


if __name__ == '__main__':
  create_question_bank()
# python -m streamlit run app.py
"""
git add .
git commit -m "UP1"
git push origin main
"""
