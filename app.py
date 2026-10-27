import datetime
import hashlib
import json
import os
import random
import re
import time
from docx import Document
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Thư mục lưu trữ hệ thống
BANK_FOLDER = 'ngan_hang_de'
SUBMIT_FOLDER = 'bai_nop'
DRAFT_FOLDER = 'ban_nhap'
IMG_FOLDER = 'cau_hoi_images'
USERS_FILE = 'users.json'
EXAMS_FILE = 'exams.json'

os.makedirs(BANK_FOLDER, exist_ok=True)
os.makedirs(SUBMIT_FOLDER, exist_ok=True)
os.makedirs(DRAFT_FOLDER, exist_ok=True)
os.makedirs(IMG_FOLDER, exist_ok=True)

# MÔN HỌC ÁP DỤNG HÌNH THỨC "TRẢ LỜI NGẮN" VÀ CHẤM TỰ ĐỘNG
SHORT_ANSWER_SUBJECTS = ['Toán', 'Vật Lý', 'Hóa Học']

# BẢNG TIỀN TỐ MÃ ĐỀ THI THEO MÔN HỌC
SUBJECT_PREFIXES = {
    'Toán': 'TO',
    'Ngữ Văn': 'NV',
    'Vật Lý': 'LY',
    'Hóa Học': 'HO',
    'Sinh Học': 'SH',
    'Tiếng Anh': 'TA',
    'Lịch Sử': 'LS',
    'Địa Lý': 'DL',
    'Tin Học': 'TH',
    'GDKTPL': 'KT',
}


# --- HÀM SO SÁNH ĐÁP ÁN TRẢ LỜI NGẮN (CHẤP NHẬN CẢ DẤU . VÀ DẤU ,) ---
def check_short_answer(user_ans, correct_ans):
  if user_ans is None or correct_ans is None:
    return False

  u_str = str(user_ans).strip().replace(' ', '')
  c_str = str(correct_ans).strip().replace(' ', '')

  if not u_str or not c_str:
    return False

  u_norm = u_str.replace(',', '.')
  c_norm = c_str.replace(',', '.')

  try:
    u_num = float(u_norm)
    c_num = float(c_norm)
    return abs(u_num - c_num) < 1e-6
  except ValueError:
    pass

  return u_norm.lower() == c_norm.lower()


# --- THUẬT TOÁN RÚT ĐỀ PHÂN BỐ ĐỀU THEO MA TRẬN CHƯƠNG & MỨC ĐỘ ---
def sample_questions_balanced(question_pool, total_needed):
  if not question_pool or total_needed <= 0:
    return []
  if len(question_pool) <= total_needed:
    selected = question_pool.copy()
    random.shuffle(selected)
    return selected

  groups = {}
  for q in question_pool:
    key = (
        q.get('chapter', 'Chương 1'),
        q.get('level', 'Thông hiểu'),
    )
    groups.setdefault(key, []).append(q)

  for key in groups:
    random.shuffle(groups[key])

  group_keys = list(groups.keys())
  random.shuffle(group_keys)

  selected = []
  while len(selected) < total_needed:
    added = False
    for key in group_keys:
      if groups[key] and len(selected) < total_needed:
        selected.append(groups[key].pop())
        added = True
    if not added:
      break

  return selected


# --- KIỂM TRA ĐỊNH DẠNG MẬT KHẨU ---
def validate_password(password):
  if len(password) < 8:
    return False, '❌ Mật khẩu phải có từ 8 ký tự trở lên!'
  if not re.search(r'[a-zA-Z]', password) or not re.search(r'[0-9]', password):
    return (
        False,
        '❌ Mật khẩu phải chứa ít nhất 1 chữ cái (a-z) và 1 chữ số (0-9)!',
    )
  return True, ''


def parse_timestamp(ts_str):
  if not ts_str:
    return datetime.datetime.now()
  for fmt in (
      '%d/%m/%Y %H:%M:%S',
      '%Y-%m-%d %H:%M:%S',
      '%d/%m/%Y %H:%M',
      '%Y-%m-%d %H:%M',
  ):
    try:
      return datetime.datetime.strptime(ts_str, fmt)
    except ValueError:
      pass
  return datetime.datetime.now()


def generate_exam_code(subject, exams_db):
  prefix = SUBJECT_PREFIXES.get(subject, 'DE')
  max_num = 0
  for ex in exams_db.values():
    if ex.get('subject') == subject:
      code = ex.get('code', '')
      if code.startswith(prefix):
        num_part = code[len(prefix) :]
        if num_part.isdigit():
          max_num = max(max_num, int(num_part))
  return f'{prefix}{max_num + 1:02d}'


def get_real_exam_code(submission_data, exams_db):
  code = submission_data.get('exam_code')
  subj = submission_data.get('subject')
  title = submission_data.get('exam_title')

  for ex in exams_db.values():
    if (
        ex.get('subject') == subj
        and ex.get('title') == title
        and ex.get('code')
    ):
      return ex.get('code')

  if code and code != 'DE':
    prefix = SUBJECT_PREFIXES.get(subj, '')
    if prefix and code.startswith(prefix):
      return code

  if code and code != 'DE':
    return code

  return 'DE'


# --- QUẢN LÝ DỮ LIỆU TÀI KHOẢN & BÀI THI ---
def load_users():
  if not os.path.exists(USERS_FILE):
    default_users = {
        'admin': {
            'name': 'Giáo Viên Quản Lý',
            'password': '123',
            'role': 'admin',
            'is_head': True,
            'subject': 'ALL',
            'managed_classes': ['ALL'],
            'contact': 'admin@school.edu.vn',
        },
        'gv_van1': {
            'name': 'Cô Phạm Thị Văn (Trưởng bộ môn)',
            'password': '123',
            'role': 'teacher',
            'is_head': True,
            'subject': 'Ngữ Văn',
            'managed_classes': ['10A1', '10A2'],
            'contact': 'vanpt@school.edu.vn',
        },
        'HS10A1_1': {
            'name': 'Nguyễn Văn A1',
            'password': '123',
            'role': 'student',
            'class_name': '10A1',
            'contact': '0900000001',
        },
    }
    save_users(default_users)
    return default_users
  with open(USERS_FILE, 'r', encoding='utf-8') as f:
    return json.load(f)


def save_users(users):
  with open(USERS_FILE, 'w', encoding='utf-8') as f:
    json.dump(users, f, ensure_ascii=False, indent=4)


def load_exams():
  if not os.path.exists(EXAMS_FILE):
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    end_str = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime(
        '%Y-%m-%d %H:%M'
    )
    default_exams = {
        'EXAM_01': {
            'code': 'NV01',
            'title': 'Bài Thi Giữa Kỳ Ngữ Văn',
            'exam_type': 'Giữa kỳ',
            'subject': 'Ngữ Văn',
            'target_class': '10A1',
            'bank_file': 'Ngan_Hang_Cau_Hoi.docx',
            'num_mcq': 10,
            'num_tf': 0,
            'num_essay': 1,
            'duration': 45,
            'max_attempts': 1,
            'mcq_points': 3.0,
            'tf_points': 0.0,
            'essay_points': 7.0,
            'start_time': now_str,
            'end_time': end_str,
            'is_active': True,
            'created_by': 'gv_van1',
        }
    }
    save_exams(default_exams)
    return default_exams
  with open(EXAMS_FILE, 'r', encoding='utf-8') as f:
    return json.load(f)


def save_exams(exams):
  with open(EXAMS_FILE, 'w', encoding='utf-8') as f:
    json.dump(exams, f, ensure_ascii=False, indent=4)


def get_student_exam_submission(student_id, exam_code):
  if not os.path.exists(SUBMIT_FOLDER):
    return None
  submissions = []
  for f_name in os.listdir(SUBMIT_FOLDER):
    if f_name.endswith('.json'):
      try:
        with open(
            os.path.join(SUBMIT_FOLDER, f_name), 'r', encoding='utf-8'
        ) as f:
          data = json.load(f)
          if (
              data.get('student_id') == student_id
              and data.get('exam_code') == exam_code
          ):
            submissions.append(data)
      except Exception:
        pass
  if not submissions:
    return None
  return sorted(submissions, key=lambda x: x.get('timestamp', ''))[-1]


def save_student_draft(
    username,
    exam_code,
    exam_info,
    start_time,
    selected_mcq,
    selected_tf,
    selected_essay,
    user_mcq,
    user_tf,
    user_essay,
    mcq_opts_map=None,
):
  draft_path = os.path.join(DRAFT_FOLDER, f'{username}_{exam_code}.json')
  draft_data = {
      'username': username,
      'exam_code': exam_code,
      'exam_info': exam_info,
      'start_time': start_time,
      'selected_mcq': selected_mcq,
      'selected_tf': selected_tf,
      'selected_essay': selected_essay,
      'mcq_opts_map': mcq_opts_map or {},
      'user_mcq': user_mcq,
      'user_tf': user_tf,
      'user_essay': user_essay,
      'updated_at': datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
  }
  with open(draft_path, 'w', encoding='utf-8') as f:
    json.dump(draft_data, f, ensure_ascii=False, indent=4)


def load_student_draft(username, exam_code):
  draft_path = os.path.join(DRAFT_FOLDER, f'{username}_{exam_code}.json')
  if os.path.exists(draft_path):
    try:
      with open(draft_path, 'r', encoding='utf-8') as f:
        return json.load(f)
    except Exception:
      return None
  return None


def delete_student_draft(username, exam_code):
  draft_path = os.path.join(DRAFT_FOLDER, f'{username}_{exam_code}.json')
  if os.path.exists(draft_path):
    try:
      os.remove(draft_path)
    except Exception:
      pass


def clear_exam_session_state():
  keys_to_del = [
      k
      for k in st.session_state.keys()
      if k.startswith('mcq_') or k.startswith('tf_') or k.startswith('essay_')
  ]
  for k in keys_to_del:
    del st.session_state[k]
  if 'exam_data' in st.session_state:
    del st.session_state['exam_data']


def restore_exam_from_draft(draft):
  """Khôi phục chính xác toàn bộ lựa chọn & câu văn bài thi từ bản nháp."""
  clear_exam_session_state()

  st.session_state.exam_data = {
      'mcq': draft['selected_mcq'],
      'tf': draft['selected_tf'],
      'essay': draft['selected_essay'],
  }
  st.session_state.start_time = draft['start_time']
  st.session_state.current_exam_info = draft['exam_info']

  mcq_opts_map = draft.get('mcq_opts_map', {})
  for idx, q in enumerate(draft['selected_mcq'], 1):
    q_id_str = str(q['id'])
    if q_id_str in mcq_opts_map:
      st.session_state[f'mcq_opts_{q["id"]}'] = mcq_opts_map[q_id_str][
          'fmt_opts'
      ]
      st.session_state[f'mcq_corr_{q["id"]}'] = mcq_opts_map[q_id_str][
          'corr_lbl'
      ]
    else:
      fmt_opts, corr_lbl = process_mcq_question(q)
      st.session_state[f'mcq_opts_{q["id"]}'] = fmt_opts
      st.session_state[f'mcq_corr_{q["id"]}'] = corr_lbl

    u_mcq = draft.get('user_mcq', {})
    val = u_mcq.get(str(idx)) or u_mcq.get(idx)
    if val:
      st.session_state[f'mcq_{idx}'] = val

  u_tf = draft.get('user_tf', {})
  for idx_key, sub_ans_list in u_tf.items():
    for sub_idx, ans_val in enumerate(sub_ans_list):
      if ans_val is not None:
        st.session_state[f'tf_{idx_key}_{sub_idx}'] = ans_val

  u_essay = draft.get('user_essay', {})
  for idx_key, essay_val in u_essay.items():
    if essay_val is not None and str(essay_val).strip():
      st.session_state[f'essay_{idx_key}'] = essay_val

  st.session_state.exam_started = True
  st.session_state.confirm_exam_id = None


# --- ĐỌC CÔNG THỨC TOÁN, HÌNH ẢNH & THẺ MA TRẬN TỪ WORD ---
def get_para_text(p):
  nodes = p._element.xpath('.//w:t | .//m:t')
  return ''.join([node.text for node in nodes if node.text]).strip()


def is_red_text(run):
  if run.font.color and run.font.color.rgb:
    rgb_str = str(run.font.color.rgb).upper()
    return rgb_str.startswith('FF') or rgb_str in [
        'E53E3E',
        'RED',
        'CC0000',
        'C00000',
    ]
  return False


def extract_paragraph_images(p, doc, q_id):
  img_paths = []
  blips = p._element.xpath('.//a:blip')
  for idx, blip in enumerate(blips):
    rId = blip.attrib.get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    )
    if rId and rId in doc.part.related_parts:
      image_part = doc.part.related_parts[rId]
      ext = image_part.content_type.split('/')[-1]
      if ext == 'jpeg':
        ext = 'jpg'

      img_filename = f'q_{q_id}_{idx}_{int(time.time()*1000)}.{ext}'
      save_path = os.path.join(IMG_FOLDER, img_filename)

      with open(save_path, 'wb') as f:
        f.write(image_part.blob)
      img_paths.append(save_path)
  return img_paths


def extract_question_metadata(raw_text):
  chapter = 'Chương 1'
  level = 'Thông hiểu'

  chap_match = re.search(
      r'\[(?:Chương|Bài|Chủ đề)\s*([\w\d\s\.-]+)\]', raw_text, re.IGNORECASE
  )
  if chap_match:
    chapter = f'Chương {chap_match.group(1).strip()}'

  if re.search(r'\[(Nhận biết|NB)\]', raw_text, re.IGNORECASE):
    level = 'Nhận biết'
  elif re.search(r'\[(Thông hiểu|TH)\]', raw_text, re.IGNORECASE):
    level = 'Thông hiểu'
  elif re.search(r'\[(Vận dụng cao|VDC)\]', raw_text, re.IGNORECASE):
    level = 'Vận dụng cao'
  elif re.search(r'\[(Vận dụng|VD)\]', raw_text, re.IGNORECASE):
    level = 'Vận dụng'

  clean_q = re.sub(
      r'\[(?:Chương|Bài|Chủ đề)\s*[\w\d\s\.-]+\]',
      '',
      raw_text,
      flags=re.IGNORECASE,
  )
  clean_q = re.sub(
      r'\[(Nhận biết|Thông hiểu|Vận dụng cao|Vận dụng|NB|TH|VD|VDC)\]',
      '',
      clean_q,
      flags=re.IGNORECASE,
  ).strip()

  return clean_q, chapter, level


@st.cache_data(show_spinner='Đang tải câu hỏi và ma trận từ Ngân hàng đề...')
def parse_question_bank(file_path):
  if not os.path.exists(file_path):
    return [], [], []
  doc = Document(file_path)
  mcq_pool, tf_pool, essay_pool = [], [], []
  current_section = None
  current_item = None

  for p in doc.paragraphs:
    text = get_para_text(p)
    has_red = any(is_red_text(run) for run in p.runs)
    is_marked = text.startswith('*') or has_red
    clean_text = text.lstrip('*').strip() if text else ''

    is_heading_1 = p.style.name.startswith('Heading 1')
    text_upper = clean_text.upper()

    if is_heading_1 or any(
        text_upper.startswith(prefix)
        for prefix in ['I.', 'II.', 'III.', 'PHẦN']
    ):
      if current_item:
        if current_section == 'MCQ':
          mcq_pool.append(current_item)
        elif current_section == 'TF':
          tf_pool.append(current_item)
        elif current_section == 'ESSAY':
          essay_pool.append(current_item)
        current_item = None

      if 'TRẮC NGHIỆM' in text_upper:
        current_section = 'MCQ'
        continue
      elif 'ĐÚNG SAI' in text_upper:
        current_section = 'TF'
        continue
      elif 'TỰ LUẬN' in text_upper or 'TRẢ LỜI NGẮN' in text_upper:
        current_section = 'ESSAY'
        continue

    if not clean_text and not p._element.xpath('.//a:blip'):
      continue

    is_q_header = re.match(
        r'^(Câu|Bài)\s*\d+[\.\:\s]?', clean_text, re.IGNORECASE
    )

    if current_section == 'MCQ':
      if is_q_header:
        if current_item:
          mcq_pool.append(current_item)
        q_id = len(mcq_pool) + 1
        raw_q = re.sub(
            r'^(Câu|Bài)\s*\d+[\.\:\s]*', '', clean_text, flags=re.IGNORECASE
        ).strip()
        q_text, chapter, level = extract_question_metadata(raw_q)
        p_imgs = extract_paragraph_images(p, doc, f'mcq_{q_id}')

        current_item = {
            'id': q_id,
            'question': q_text,
            'options': [],
            'images': p_imgs,
            'chapter': chapter,
            'level': level,
        }
      elif current_item:
        p_imgs = extract_paragraph_images(
            p, doc, f"mcq_{current_item['id']}_sub"
        )
        if p_imgs:
          current_item['images'].extend(p_imgs)

        if re.match(r'^[A-D][\.\:\)]', clean_text, re.IGNORECASE):
          opt_text = re.sub(
              r'^[A-D][\.\:\)]\s*', '', clean_text, flags=re.IGNORECASE
          ).strip()
          current_item['options'].append(
              {'text': opt_text, 'is_correct': is_marked}
          )
        elif clean_text and not current_item['options']:
          current_item['question'] += ' ' + clean_text

    elif current_section == 'TF':
      if is_q_header:
        if current_item:
          tf_pool.append(current_item)
        q_id = len(tf_pool) + 1
        raw_q = re.sub(
            r'^(Câu|Bài)\s*\d+[\.\:\s]*', '', clean_text, flags=re.IGNORECASE
        ).strip()
        q_text, chapter, level = extract_question_metadata(raw_q)
        p_imgs = extract_paragraph_images(p, doc, f'tf_{q_id}')

        current_item = {
            'id': q_id,
            'question': q_text,
            'sub_items': [],
            'images': p_imgs,
            'chapter': chapter,
            'level': level,
        }
      elif current_item:
        p_imgs = extract_paragraph_images(p, doc, f"tf_{current_item['id']}_sub")
        if p_imgs:
          current_item['images'].extend(p_imgs)

        if re.match(r'^[a-d][\)\:\.]', clean_text, re.IGNORECASE):
          is_true = is_marked or '[ĐÚNG]' in clean_text.upper()
          clean_item = re.sub(
              r'\[(ĐÚNG|SAI)\]', '', clean_text, flags=re.IGNORECASE
          ).strip()
          clean_item = re.sub(
              r'^[a-d][\)\:\.]\s*', '', clean_item, flags=re.IGNORECASE
          ).strip()
          current_item['sub_items'].append(
              {'text': clean_item, 'answer': is_true}
          )
        elif clean_text and not current_item['sub_items']:
          current_item['question'] += ' ' + clean_text

    elif current_section == 'ESSAY':
      is_essay_header = (
          is_q_header
          and not clean_text.lower().startswith('đáp án')
          and not clean_text.lower().startswith('gợi ý')
      )
      if is_essay_header:
        if current_item:
          essay_pool.append(current_item)
        q_id = len(essay_pool) + 1
        raw_q = re.sub(
            r'^(Câu|Bài)\s*\d+[\.\:\s]*', '', clean_text, flags=re.IGNORECASE
        ).strip()
        q_text, chapter, level = extract_question_metadata(raw_q)
        p_imgs = extract_paragraph_images(p, doc, f'es_{q_id}')

        current_item = {
            'id': q_id,
            'question': q_text,
            'sample_answer': '',
            'images': p_imgs,
            'chapter': chapter,
            'level': level,
        }
      elif current_item:
        p_imgs = extract_paragraph_images(p, doc, f"es_{current_item['id']}_sub")
        if p_imgs:
          current_item['images'].extend(p_imgs)

        if text.startswith('*'):
          current_item['sample_answer'] = text.lstrip('*').strip()
        elif clean_text.lower().startswith(
            'đáp án'
        ) or clean_text.lower().startswith('gợi ý'):
          current_item['sample_answer'] = clean_text.split(':', 1)[-1].strip()
        elif clean_text:
          current_item['question'] += ' ' + clean_text

  if current_item:
    if current_section == 'MCQ':
      mcq_pool.append(current_item)
    elif current_section == 'TF':
      tf_pool.append(current_item)
    elif current_section == 'ESSAY':
      essay_pool.append(current_item)

  return mcq_pool, tf_pool, essay_pool


def process_mcq_question(q):
  labels = ['A', 'B', 'C', 'D']
  opts = q['options'].copy()
  random.shuffle(opts)

  formatted_options = []
  correct_label = 'A'

  for idx, opt in enumerate(opts):
    lbl = labels[idx] if idx < len(labels) else f'Opt{idx+1}'
    formatted_options.append(f'{lbl}. {opt["text"]}')
    if opt['is_correct']:
      correct_label = lbl

  return formatted_options, correct_label


def grade_tf(user_answers, correct_answers):
  correct_count = sum(
      1 for u, c in zip(user_answers, correct_answers) if u == c
  )
  score_map = {0: 0.0, 1: 0.1, 2: 0.25, 3: 0.5, 4: 1.0}
  return score_map.get(correct_count, 0.0), correct_count


# --- GIAO DIỆN & DỊCH VỤ NỀN ---
st.set_page_config(page_title='Hệ Thống Thi Thử Trực Tuyến', layout='wide')

if 'bank_uploader_key' not in st.session_state:
  st.session_state.bank_uploader_key = 0
if 'restore_uploader_key' not in st.session_state:
  st.session_state.restore_uploader_key = 0
if 'excel_uploader_key' not in st.session_state:
  st.session_state.excel_uploader_key = 0

st.markdown(
    """
    <style>
    html, body, p, label, input, textarea { font-family: 'Arial', sans-serif !important; }
    button:not([aria-label*="password"]):not([aria-label*="Password"]) { font-family: 'Arial', sans-serif !important; }
    h3, .stSubheader, div[data-testid="stSubheader"] *, div[data-testid="stMarkdownContainer"] h3 {
        font-family: 'Arial', sans-serif !important;
        font-size: 17px !important;
        font-weight: bold !important;
        color: #1f1f1f !important;
        margin-bottom: 8px !important;
    }
    [data-testid="stIcon"], .material-symbols-rounded, [class*="stIcon"], i[class*="st-"], button[aria-label*="password"] *, button[aria-label*="Password"] * {
        font-family: 'Material Symbols Rounded', 'Material Icons' !important;
    }
    div[id^="q_"] { scroll-margin-top: 70px !important; display: block; }
    div[role="radiogroup"] label p, label p { font-family: 'Arial', sans-serif !important; font-size: 17px !important; color: #2b2b2b !important; }
    textarea, input[type="text"] { font-family: 'Arial', sans-serif !important; font-size: 17px !important; }
    
    /* ẨN POPUP MẬT KHẨU BẰNG CÁCH MÃ HÓA Ô NHẬP TỰ ĐỘNG THÀNH DẤU CHẤM */
    div[data-testid="stTextInput"] input[aria-label*="Mật khẩu"] {
        -webkit-text-security: disc !important;
    }

    /* Cấu hình khung đăng nhập chuẩn đẹp */
    .login-box { 
        background-color: #ffffff; 
        padding: 35px 30px; 
        border-radius: 14px; 
        box-shadow: 0px 4px 20px rgba(0,0,0,0.12); 
        margin: 40px auto; 
        max-width: 500px;
        width: 100%;
    }
    .login-title {
        font-size: 24px !important;
        font-weight: bold !important;
        text-align: center !important;
        white-space: nowrap !important;
        color: #1f1f1f !important;
        margin-bottom: 25px !important;
    }
    .confirm-modal { background-color: #ffffff; padding: 30px; border-radius: 14px; border: 2px solid #007bff; box-shadow: 0px 10px 30px rgba(0,0,0,0.2); width: 100%; }
    </style>
""",
    unsafe_allow_html=True,
)

# JavaScript tự động đăng xuất sau 5 phút không thao tác
inactivity_js = """
<script>
(function() {
    var timer;
    function resetTimer() {
        clearTimeout(timer);
        timer = setTimeout(function() { window.parent.location.reload(); }, 300000); 
    }
    try {
        window.parent.document.onmousemove = resetTimer;
        window.parent.document.onkeypress = resetTimer;
        window.parent.document.onclick = resetTimer;
        window.parent.document.onscroll = resetTimer;
        resetTimer();
    } catch(e) {}
})();
</script>
"""
components.html(inactivity_js, height=0, width=0)

users_db = load_users()
exams_db = load_exams()

if 'user_session' not in st.session_state:
  st.session_state.user_session = None
if 'show_forgot' not in st.session_state:
  st.session_state.show_forgot = False
if 'submitted_data' not in st.session_state:
  st.session_state.submitted_data = None
if 'confirm_exam_id' not in st.session_state:
  st.session_state.confirm_exam_id = None

now_ts = time.time()
if st.session_state.user_session:
  last_act = st.session_state.get('last_activity', now_ts)
  if now_ts - last_act > 300:
    clear_exam_session_state()
    st.session_state.user_session = None
    st.session_state.submitted_data = None
    st.session_state.confirm_exam_id = None
    st.warning('⏰ Đã tự động đăng xuất do bạn không hoạt động trong 5 phút!')
    st.rerun()
  else:
    st.session_state.last_activity = now_ts


# --- MÀN HÌNH ĐĂNG NHẬP ---
if not st.session_state.user_session:
  col_l1, col_center, col_l2 = st.columns([1, 2, 1])
  with col_center:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.markdown(
        "<div class='login-title'>🔐 ĐĂNG NHẬP HỆ THỐNG</div>",
        unsafe_allow_html=True,
    )

    username_input = st.text_input('Mã học sinh / Tên đăng nhập:')

    # Bỏ type='password' để triệt tiêu popup gợi ý mật khẩu mạnh của Chrome, sử dụng CSS -webkit-text-security thay thế
    password_input = st.text_input('Mật khẩu:')

    if st.button('🚀 Đăng Nhập', use_container_width=True):
      if username_input in users_db:
        if users_db[username_input]['password'] == password_input:
          u = users_db[username_input]
          st.session_state.user_session = {
              'username': username_input,
              'name': u['name'],
              'role': u['role'],
              'is_head': u.get('is_head', False),
              'subject': u.get('subject', 'ALL'),
              'managed_classes': u.get('managed_classes', ['ALL']),
              'class_name': u.get('class_name', ''),
          }
          st.session_state.last_activity = time.time()
          st.success('Đăng nhập thành công!')
          st.rerun()
        else:
          st.error('❌ Mật khẩu không chính xác!')
      else:
        st.error('❌ Mã đăng nhập không tồn tại!')

    if st.button('Quên mật khẩu?', key='btn_forgot_toggle'):
      st.session_state.show_forgot = not st.session_state.show_forgot

    if st.session_state.show_forgot:
      st.markdown('---')
      st.subheader('🔑 Khôi phục mật khẩu')
      fg_user = st.text_input('Nhập Mã học sinh / Tên đăng nhập:')
      fg_contact = st.text_input('Nhập Email/SĐT đã đăng ký:')
      if st.button('Xác thực thông tin'):
        if (
            fg_user in users_db
            and users_db[fg_user].get('contact') == fg_contact
        ):
          st.session_state.reset_user = fg_user
          st.success('✅ Xác thực thành công! Hãy nhập mật khẩu mới bên dưới.')
        else:
          st.error('❌ Thông tin đăng ký không chính xác!')

      if st.session_state.get('reset_user'):
        new_pass = st.text_input('Nhập mật khẩu mới:')
        confirm_pass = st.text_input('Nhập lại mật khẩu mới:')
        if st.button('✅ Đổi Mật Khẩu'):
          if new_pass != confirm_pass:
            st.error('❌ Mật khẩu xác nhận không khớp!')
          else:
            is_valid, err_msg = validate_password(new_pass)
            if not is_valid:
              st.error(err_msg)
            else:
              u_target = st.session_state.reset_user
              users_db[u_target]['password'] = new_pass
              save_users(users_db)
              st.success('Đổi mật khẩu thành công! Bạn có thể đăng nhập ngay.')
              st.session_state.show_forgot = False
              st.session_state.reset_user = None

    st.markdown('</div>', unsafe_allow_html=True)
  st.stop()


# --- HEADER BÊN PHẢI ---
user = st.session_state.user_session
head_col1, head_col2 = st.columns([3, 1])
with head_col2:
  role_display = (
      user['class_name']
      if user['role'] == 'student'
      else (f"GV {user['subject']}" if user['role'] == 'teacher' else 'ADMIN')
  )
  with st.popover(f"👤 {user['name']} ({role_display})"):
    st.write(f"**Tài khoản:** `{user['username']}`")
    st.write(
        f"**Chức vụ:**"
        f" {'Trưởng bộ môn' if user.get('is_head') else user['role'].upper()}"
    )

    with st.expander('🔑 Đổi Mật Khẩu'):
      with st.form(key=f"form_change_pass_{user['username']}"):
        old_p = st.text_input('Mật khẩu cũ:')
        new_p = st.text_input('Mật khẩu mới:')
        confirm_p = st.text_input('Nhập lại mật khẩu mới:')
        sub_pass = st.form_submit_button('💾 Lưu Mật Khẩu Mới')

        if sub_pass:
          current_pass = users_db[user['username']]['password']
          if old_p != current_pass:
            st.error('❌ Mật khẩu cũ không chính xác!')
          elif new_p != confirm_p:
            st.error('❌ Mật khẩu mới và Nhập lại không khớp!')
          else:
            is_valid, err_msg = validate_password(new_p)
            if not is_valid:
              st.error(err_msg)
            else:
              users_db[user['username']]['password'] = new_p
              save_users(users_db)
              st.success('✅ Đã đổi mật khẩu thành công!')

    if st.button('🚪 Đăng Xuất', use_container_width=True):
      clear_exam_session_state()
      st.session_state.user_session = None
      st.session_state.submitted_data = None
      st.session_state.confirm_exam_id = None
      if 'exam_started' in st.session_state:
        del st.session_state['exam_started']
      st.rerun()


# ==================== GIAO DIỆN GIÁO VIÊN / ADMIN ====================
if user['role'] in ['teacher', 'admin']:
  st.title('👨‍🏫 CỔNG QUẢN LÝ THI THỬ')

  tabs_list = [
      '📝 Chấm Bài Thi',
      '📊 Thống Kê Điểm',
      '⚙️ Tạo & Quản Lý Đề Thi',
      '🏫 Quản Lý Lớp Học',
  ]
  if user['role'] == 'admin' or user.get('is_head'):
    tabs_list.append('👨‍🏫 Quản Lý Giáo Viên & Phân Công')

  tabs = st.tabs(tabs_list)

  # TAB 1: CHẤM BÀI THI
  with tabs[0]:
    st.subheader('📝 BẢNG CHẤM BÀI THI CỦA HỌC SINH')
    submission_files = [
        f for f in os.listdir(SUBMIT_FOLDER) if f.endswith('.json')
    ]
    raw_subs = []

    for f_name in submission_files:
      try:
        with open(
            os.path.join(SUBMIT_FOLDER, f_name), 'r', encoding='utf-8'
        ) as f:
          data = json.load(f)
          st_id = data.get('student_id', '')
          st_class = users_db.get(st_id, {}).get(
              'class_name', data.get('class_name', '10A1')
          )
          data['class_name'] = st_class
          data['file_name'] = f_name
          data['real_code'] = get_real_exam_code(data, exams_db)

          if user['role'] == 'admin':
            raw_subs.append(data)
          else:
            ex_type = data.get('exam_type', 'Thường xuyên')
            if ex_type == 'Cuối kỳ':
              if data.get('assigned_teacher') == user['username']:
                raw_subs.append(data)
            else:
              if (
                  data.get('created_by') == user['username']
                  or data.get('subject') == user['subject']
              ):
                raw_subs.append(data)
      except Exception:
        pass

    if not raw_subs:
      st.info('Hiện tại không có bài thi nào thuộc quyền chấm của bạn.')
    else:
      if user['role'] == 'admin':
        avail_subjects = sorted(
            list(set(d.get('subject', '') for d in raw_subs if d.get('subject')))
        )
        sel_subject = st.selectbox(
            '1️⃣ Chọn Môn Học:', ['-- Chọn Môn Học --'] + avail_subjects
        )
      else:
        sel_subject = user['subject']
        st.info(f'📚 **Môn học mặc định của giáo viên:** {sel_subject}')

      if sel_subject and sel_subject != '-- Chọn Môn Học --':
        subs_by_sub = [d for d in raw_subs if d.get('subject') == sel_subject]
        if not subs_by_sub:
          st.warning(f'Chưa có bài nộp nào thuộc môn {sel_subject}.')
        else:
          avail_classes = sorted(
              list(
                  set(
                      d.get('class_name', '')
                      for d in subs_by_sub
                      if d.get('class_name')
                  )
              )
          )
          sel_class = st.selectbox(
              '2️⃣ Chọn Lớp Học:', ['-- Chọn Lớp Học --'] + avail_classes
          )

          if sel_class and sel_class != '-- Chọn Lớp Học --':
            candidates = [
                d for d in subs_by_sub if d.get('class_name') == sel_class
            ]

            if not candidates:
              st.warning(f'Chưa có bài nộp nào thuộc lớp {sel_class}.')
            else:
              candidates.sort(
                  key=lambda x: (
                      1 if x.get('is_graded', False) else 0,
                      x.get('timestamp', ''),
                  )
              )

              sub_options_map = {}
              options_list = ['-- Chọn Học Sinh / Bài Thi --']

              for d in candidates:
                f_name = d['file_name']
                is_g = d.get('is_graded', False)
                status_tag = '✅ [ĐÃ CHẤM]' if is_g else '⏳ [CHƯA CHẤM]'

                is_anon = (
                    d.get('exam_type') == 'Cuối kỳ' and user['role'] != 'admin'
                )
                if is_anon:
                  anon_code = hashlib.md5(
                      d.get('student_id', '').encode()
                  ).hexdigest()[:6].upper()
                  st_disp = f'Thí sinh #{anon_code}'
                else:
                  st_disp = (
                      f"HS: {d.get('student_name')} ({d.get('student_id')})"
                  )

                score_tag = f" - {d.get('total_score', 0)}đ" if is_g else ''
                label = f"{status_tag} [{d['real_code']}] {d.get('exam_title')} - {st_disp}{score_tag}"

                sub_options_map[label] = d
                options_list.append(label)

              selected_sub_label = st.selectbox(
                  '3️⃣ Chọn Bài Nộp Cần Chấm:', options_list
              )

              if (
                  selected_sub_label
                  and selected_sub_label != '-- Chọn Học Sinh / Bài Thi --'
              ):
                data = sub_options_map[selected_sub_label]
                f_name = data['file_name']
                sub_path = os.path.join(SUBMIT_FOLDER, f_name)

                real_code = data['real_code']
                is_final_exam = data.get('exam_type') == 'Cuối kỳ'
                is_graded = data.get('is_graded', False)

                sub_dt = parse_timestamp(data.get('timestamp', ''))
                now_dt = datetime.datetime.now()
                days_diff = (now_dt - sub_dt).total_seconds() / (24 * 3600)
                can_edit = days_diff <= 14

                st.markdown('---')

                if is_final_exam and user['role'] != 'admin':
                  anon_code = hashlib.md5(
                      data.get('student_id', '').encode()
                  ).hexdigest()[:6].upper()
                  st.subheader(
                      f'🔒 BÀI THI CUỐI KỲ ẨN DANH: Thí sinh #{anon_code} (Môn:'
                      f' {data.get("subject")})'
                  )
                else:
                  st.subheader(
                      f'Học sinh: {data.get("student_name")} - Mã HS:'
                      f' {data.get("student_id")} (Lớp: {data.get("class_name")})'
                  )

                st.caption(
                    f'Bài thi: [{real_code}] {data.get("exam_title")}'
                    f' ({data.get("exam_type")}) | Thời gian nộp:'
                    f' {data.get("timestamp")}'
                )

                edit_mode = False
                if is_graded:
                  col_st1, col_st2 = st.columns([3, 1])
                  with col_st1:
                    st.success(
                        f'✅ Bài thi này **ĐÃ CHẤM** (Tổng điểm chốt:'
                        f" **{data.get('total_score', 0.0)}đ**)"
                    )
                  with col_st2:
                    if can_edit:
                      edit_key = f'edit_btn_{f_name}'
                      if edit_key not in st.session_state:
                        st.session_state[edit_key] = False
                      if st.button(
                          '✏️ Sửa Điểm', key=f'btn_toggle_edit_{f_name}'
                      ):
                        st.session_state[edit_key] = not st.session_state[
                            edit_key
                        ]
                        st.rerun()
                      edit_mode = st.session_state[edit_key]
                    else:
                      st.error('🔒 Quá 14 ngày (Không thể sửa)')

                  if edit_mode:
                    st.info('💡 Bạn đang ở chế độ chỉnh sửa điểm bài đã chấm.')
                else:
                  edit_mode = True

                max_essay_points = float(data.get('essay_points', 0.0))
                new_essay_scores = []

                if data.get('essay_details'):
                  section_title = (
                      'Trả lời ngắn'
                      if data.get('subject') in SHORT_ANSWER_SUBJECTS
                      else 'Tự luận'
                  )
                  st.header(
                      f'Chấm điểm phần {section_title} (Tổng tối đa:'
                      f' {max_essay_points}đ)'
                  )
                  for idx, item in enumerate(data['essay_details'], 1):
                    st.write(f'**Câu {idx}:** {item["question"]}')
                    st.write(
                        f'**Bài làm thí sinh:** `{item["user_text"] or "Bỏ trống"}`'
                    )
                    st.write(f'**Đáp án chuẩn:** `{item["sample_answer"]}`')

                    if is_graded and not edit_mode:
                      st.write(
                          '**Điểm đã chấm:**'
                          f' `{item.get("final_score", 0.0)}đ`'
                      )
                    else:
                      new_sc = st.number_input(
                          f'Chốt điểm câu {idx}:',
                          0.0,
                          max_essay_points,
                          float(item.get('final_score', 0.0)),
                          step=0.25,
                          key=f're_{f_name}_{idx}',
                          disabled=not (can_edit or not is_graded),
                      )
                      new_essay_scores.append(new_sc)

                if is_graded and not edit_mode:
                  st.text_area(
                      'Nhận xét bài làm:',
                      value=data.get('teacher_note', ''),
                      disabled=True,
                  )
                else:
                  teacher_note = st.text_area(
                      'Nhận xét bài làm:',
                      value=data.get('teacher_note', ''),
                      disabled=not (can_edit or not is_graded),
                  )

                  if st.button(
                      '💾 Lưu Điểm Chốt & Công Bố'
                      if not is_graded
                      else '💾 Cập Nhật Điểm Mới',
                      disabled=not (can_edit or not is_graded),
                  ):
                    sum_essay = (
                        sum(new_essay_scores) if new_essay_scores else 0.0
                    )
                    if sum_essay > max_essay_points and max_essay_points > 0:
                      st.error(
                          f'❌ Tổng điểm phần này ({sum_essay}đ) vượt quá điểm'
                          f' cho phép ({max_essay_points}đ)!'
                      )
                    else:
                      auto_score = float(data.get('auto_score', 0.0))
                      if data.get('essay_details'):
                        for idx, new_sc in enumerate(new_essay_scores):
                          data['essay_details'][idx]['final_score'] = new_sc

                      if data.get('subject') in SHORT_ANSWER_SUBJECTS:
                        data['total_score'] = round(auto_score, 2)
                      else:
                        data['total_score'] = round(auto_score + sum_essay, 2)

                      data['teacher_note'] = teacher_note
                      data['is_graded'] = True

                      with open(sub_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=4)
                      st.success(
                          '✅ Đã lưu điểm thành công! Tổng điểm chốt:'
                          f' {data["total_score"]}đ'
                      )
                      if f'edit_btn_{f_name}' in st.session_state:
                        st.session_state[f'edit_btn_{f_name}'] = False
                      st.rerun()

  # TAB 2: THỐNG KÊ ĐIỂM BÀI THI
  with tabs[1]:
    st.subheader('📊 THỐNG KÊ KẾT QUẢ ĐIỂM THI')
    all_submissions = []
    for f_name in os.listdir(SUBMIT_FOLDER):
      if f_name.endswith('.json'):
        try:
          with open(
              os.path.join(SUBMIT_FOLDER, f_name), 'r', encoding='utf-8'
          ) as f:
            data = json.load(f)
            st_id = data.get('student_id')
            st_class = (
                users_db.get(st_id, {}).get('class_name', 'N/A')
                if st_id in users_db
                else 'N/A'
            )
            data['class_name'] = st_class
            data['real_code'] = get_real_exam_code(data, exams_db)

            if user['role'] == 'admin':
              all_submissions.append(data)
            elif user['role'] == 'teacher':
              if data.get('subject') == user['subject']:
                all_submissions.append(data)
        except Exception:
          pass

    if not all_submissions:
      st.info('Chưa có dữ liệu bài nộp nào để thống kê.')
    else:
      f_col1, f_col2, f_col3 = st.columns(3)
      classes_list = ['Tất cả lớp'] + sorted(
          list(set(d['class_name'] for d in all_submissions))
      )
      subjects_list = ['Tất cả môn'] + sorted(
          list(set(d.get('subject', '') for d in all_submissions))
      )
      exams_list = ['Tất cả bài thi'] + sorted(
          list(set(d.get('real_code', '') for d in all_submissions))
      )

      with f_col1:
        sel_class = st.selectbox('Lọc theo Lớp:', classes_list)
      with f_col2:
        sel_subject = st.selectbox('Lọc theo Môn:', subjects_list)
      with f_col3:
        sel_exam = st.selectbox('Lọc theo Bài thi:', exams_list)

      filtered_subs = []
      for d in all_submissions:
        match_class = (
            sel_class == 'Tất cả lớp' or d.get('class_name') == sel_class
        )
        match_subject = (
            sel_subject == 'Tất cả môn' or d.get('subject') == sel_subject
        )
        match_exam = (
            sel_exam == 'Tất cả bài thi' or d.get('real_code') == sel_exam
        )
        if match_class and match_subject and match_exam:
          filtered_subs.append(d)

      if not filtered_subs:
        st.warning('Không tìm thấy kết quả phù hợp với bộ lọc.')
      else:
        scores = [f.get('total_score', 0.0) for f in filtered_subs]
        avg_score = round(sum(scores) / len(scores), 2)
        max_score = max(scores)
        min_score = min(scores)
        pass_rate = round(
            sum(1 for s in scores if s >= 5.0) / len(scores) * 100, 1
        )
        good_rate = round(
            sum(1 for s in scores if s >= 8.0) / len(scores) * 100, 1
        )

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric('Tổng bài nộp', f'{len(scores)} bài')
        m2.metric('Điểm trung bình', f'{avg_score}đ')
        m3.metric('Điểm Cao/Thấp nhất', f'{max_score}đ / {min_score}đ')
        m4.metric('Tỷ lệ Đạt (≥5.0)', f'{pass_rate}%')
        m5.metric('Tỷ lệ Giỏi (≥8.0)', f'{good_rate}%')

        st.markdown('---')
        st.markdown('### 📋 Danh Sách Điểm Chi Tiết')

        table_data = []
        for f in filtered_subs:
          is_graded = f.get('is_graded', True)
          is_anon = (
              f.get('exam_type') == 'Cuối kỳ' and user['role'] != 'admin'
          )
          anon_name = (
              f"Thí sinh #{hashlib.md5(f.get('student_id','').encode()).hexdigest()[:6].upper()}"
          )

          table_data.append({
              'Mã HS': '🔒 Ẩn danh' if is_anon else f.get('student_id'),
              'Họ và Tên': anon_name if is_anon else f.get('student_name'),
              'Lớp': f.get('class_name'),
              'Mã Đề': f.get('real_code'),
              'Tên Bài Thi': f.get('exam_title'),
              'Môn': f.get('subject'),
              'Điểm Tự Động': f.get('auto_score'),
              'Tổng Điểm': f.get('total_score'),
              'Trạng Thái': '🟢 Đã chấm' if is_graded else '⏳ Chờ chấm',
              'Thời Gian Nộp': f.get('timestamp'),
          })

        st.dataframe(table_data, use_container_width=True)

  # TAB 3: TẠO & QUẢN LÝ ĐỀ THI
  with tabs[2]:
    st.subheader('📤 TẢI LÊN & QUẢN LÝ NGÂN HÀNG ĐỀ THI')

    uploaded_bank = st.file_uploader(
        'Chọn file câu hỏi Word (.docx) từ máy tính (Tải lên thành công sẽ tự'
        ' động cập nhật vào danh sách):',
        type=['docx'],
        key=f'bank_uploader_{st.session_state.bank_uploader_key}',
    )

    if uploaded_bank is not None:
      saved_filename = f"{user['username']}_{uploaded_bank.name}"
      save_path = os.path.join(BANK_FOLDER, saved_filename)
      is_overwrite = os.path.exists(save_path)

      with open(save_path, 'wb') as f:
        f.write(uploaded_bank.getbuffer())

      parse_question_bank.clear()

      st.session_state.bank_uploader_key += 1

      if is_overwrite:
        st.toast(
            f'🔄 Đã **CẬP NHẬT / GHI ĐÈ** thành công file `{uploaded_bank.name}`!'
        )
      else:
        st.toast(f'✅ Đã tải lên file mới `{uploaded_bank.name}` thành công!')

      st.rerun()

    raw_bank_files = [
        f for f in os.listdir(BANK_FOLDER) if f.endswith('.docx')
    ]
    if user['role'] == 'admin':
      my_bank_files = raw_bank_files
    else:
      prefix = f"{user['username']}_"
      my_bank_files = [f for f in raw_bank_files if f.startswith(prefix)]

    if my_bank_files:
      with st.expander(
          '🗂️ Quản Lý File Ngân Hàng Đề Của Bạn (Xem / Xóa file đã up)'
      ):
        for f_item in my_bank_files:
          disp_name = (
              f_item.split('_', 1)[-1] if '_' in f_item else f_item
          )
          col_f1, col_f2 = st.columns([4, 1])
          with col_f1:
            st.write(f'📄 **{disp_name}**')
          with col_f2:
            if st.button('🗑️ Xóa File', key=f'del_file_{f_item}'):
              os.remove(os.path.join(BANK_FOLDER, f_item))
              parse_question_bank.clear()
              st.toast(f'Đã xóa file {disp_name}')
              st.rerun()

    st.markdown('---')
    st.subheader('➕ Tạo Đề Thi Mới')

    if not my_bank_files:
      st.warning(
          '💡 Bạn chưa có file ngân hàng đề thi nào. Hãy dùng nút **Upload**'
          ' bên trên để đưa file Word (.docx) của bạn vào hệ thống!'
      )
    else:
      c1, c2 = st.columns(2)
      with c1:
        if user['role'] == 'teacher':
          ex_subject = st.text_input(
              'Môn học:', value=user['subject'], disabled=True
          )
          allowed_classes = user.get('managed_classes', [])
          target_class = st.selectbox(
              'Lớp áp dụng:',
              options=allowed_classes if allowed_classes else ['ALL'],
          )
        else:
          ex_subject = st.selectbox(
              'Môn học:', options=list(SUBJECT_PREFIXES.keys())
          )
          target_class = st.text_input(
              'Lớp áp dụng (VD: 10A1 hoặc ALL):', value='10A1'
          )

        auto_code = generate_exam_code(ex_subject, exams_db)
        ex_code = st.text_input('Mã bài thi (Tự động):', value=auto_code)
        ex_title = st.text_input('Tên bài thi:', value='Bài Thi Giữa Kỳ')

        if user['role'] == 'admin' or user.get('is_head'):
          allowed_types = ['Thường xuyên', 'Giữa kỳ', 'Cuối kỳ']
        else:
          allowed_types = ['Thường xuyên', 'Giữa kỳ']

        ex_type = st.selectbox('Dạng bài thi:', options=allowed_types)
        max_attempts_input = st.number_input('Số lần làm bài tối đa:', 1, 10, 1)

      with c2:
        display_map = {
            (f.split('_', 1)[-1] if '_' in f else f): f for f in my_bank_files
        }
        selected_display = st.selectbox(
            'Chọn File Ngân hàng đề của bạn:', list(display_map.keys())
        )
        selected_bank = display_map[selected_display]

        mcq_p, tf_p, es_p = parse_question_bank(
            os.path.join(BANK_FOLDER, selected_bank)
        )

        all_pool = mcq_p + tf_p + es_p
        chap_set = set(q.get('chapter', 'Chương 1') for q in all_pool)
        level_set = set(q.get('level', 'Thông hiểu') for q in all_pool)
        st.caption(
            f'📊 **Thống kê ma trận file:** {len(all_pool)} câu |'
            f' {len(chap_set)} Chương | Các mức độ:'
            f' {", ".join(sorted(level_set))}'
        )

        num_mcq_input = st.number_input(
            'Số câu Trắc nghiệm:',
            min_value=0,
            max_value=len(mcq_p),
            value=min(10, len(mcq_p)),
        )
        num_tf_input = st.number_input(
            'Số câu Đúng/Sai:',
            min_value=0,
            max_value=len(tf_p),
            value=min(2, len(tf_p)),
        )

        essay_label = (
            'Trả lời ngắn' if ex_subject in SHORT_ANSWER_SUBJECTS else 'Tự luận'
        )
        num_es_input = st.number_input(
            f'Số câu {essay_label}:',
            min_value=0,
            max_value=len(es_p),
            value=min(2, len(es_p)),
        )
        duration_input = st.number_input('Thời gian làm bài (Phút):', 1, 180, 45)

      st.markdown('### ⏱️ Cấu Hình Thời Gian Mở Đề')
      t_col1, t_col2 = st.columns(2)
      now_dt = datetime.datetime.now()
      with t_col1:
        start_d = st.date_input('Ngày bắt đầu mở đề:', now_dt.date())
        start_t = st.time_input('Giờ bắt đầu mở đề:', now_dt.time())

      start_datetime = datetime.datetime.combine(start_d, start_t)
      default_end = start_datetime + datetime.timedelta(
          minutes=duration_input + 5
      )

      with t_col2:
        end_d = st.date_input('Ngày kết thúc mở đề:', default_end.date())
        end_t = st.time_input('Giờ kết thúc mở đề:', default_end.time())

      end_datetime = datetime.datetime.combine(end_d, end_t)

      if ex_subject == 'Ngữ Văn' and (num_mcq_input > 0 or num_tf_input > 0):
        essay_points = 7.0
        tf_points = float(num_tf_input) * 1.0
        mcq_points = max(0.0, 3.0 - tf_points)
      elif ex_subject == 'Toán':
        essay_points = round(float(num_es_input) * 0.5, 2)
        tf_points = float(num_tf_input) * 1.0
        mcq_points = max(0.0, round(10.0 - tf_points - essay_points, 2))
      elif ex_subject in ['Vật Lý', 'Hóa Học']:
        essay_points = round(float(num_es_input) * 0.25, 2)
        tf_points = float(num_tf_input) * 1.0
        mcq_points = max(0.0, round(10.0 - tf_points - essay_points, 2))
      else:
        tf_points = float(num_tf_input) * 1.0
        remaining_pts = max(0.0, 10.0 - tf_points)

        if num_mcq_input == 0 and num_es_input == 0:
          mcq_points, essay_points = 0.0, 0.0
        elif num_mcq_input == 0 and num_es_input > 0:
          mcq_points, essay_points = 0.0, remaining_pts
        elif num_mcq_input > 0 and num_es_input == 0:
          mcq_points, essay_points = remaining_pts, 0.0
        else:
          essay_points = st.number_input(
              f'Nhập tổng điểm phần {essay_label} (Tối đa {remaining_pts}đ):',
              min_value=0.0,
              max_value=remaining_pts,
              value=min(3.0, remaining_pts),
              step=0.5,
              key='input_essay_pts',
          )
          mcq_points = max(0.0, remaining_pts - essay_points)

      st.info(
          f'📊 **Cơ cấu điểm đề thi:** Trắc nghiệm: **{mcq_points}đ** | Đúng/Sai:'
          f' **{tf_points}đ** | {essay_label}: **{essay_points}đ** (Tổng: 10.0đ)'
      )

      if st.button('💾 Tạo Đề Thi'):
        if end_datetime <= start_datetime:
          st.error('❌ Thời gian kết thúc phải lớn hơn thời gian bắt đầu mở đề!')
        else:
          exam_id = f'EXAM_{int(time.time())}'
          exams_db[exam_id] = {
              'code': ex_code,
              'title': ex_title,
              'exam_type': ex_type,
              'subject': ex_subject,
              'target_class': target_class,
              'bank_file': selected_bank,
              'num_mcq': num_mcq_input,
              'num_tf': num_tf_input,
              'num_essay': num_es_input,
              'duration': duration_input,
              'max_attempts': max_attempts_input,
              'mcq_points': mcq_points,
              'tf_points': tf_points,
              'essay_points': essay_points,
              'start_time': start_datetime.strftime('%Y-%m-%d %H:%M'),
              'end_time': end_datetime.strftime('%Y-%m-%d %H:%M'),
              'is_active': True,
              'created_by': user['username'],
          }
          save_exams(exams_db)
          st.success(
              f'✅ Đã tạo thành công bài thi [{ex_code}] ({ex_type}) cho môn'
              f' {ex_subject}!'
          )
          st.rerun()

    st.markdown('---')
    st.subheader('📋 Danh Sách Đề Thi Đã Tạo')

    visible_exam_items = [
        (ex_id, ex_info)
        for ex_id, ex_info in exams_db.items()
        if user['role'] == 'admin' or ex_info.get('created_by') == user['username']
    ]

    if not visible_exam_items:
      st.info('Bạn chưa tạo đề thi nào.')
    else:
      for ex_id, ex_info in sorted(
          visible_exam_items, key=lambda x: x[1].get('code', '')
      ):
        ex_code_disp = ex_info.get('code', 'DE')
        e_lbl = (
            'Trả lời ngắn'
            if ex_info.get('subject') in SHORT_ANSWER_SUBJECTS
            else 'Tự luận'
        )
        with st.expander(
            f"📌 [{ex_code_disp}] {ex_info['title']}"
            f" ({ex_info.get('exam_type','Thường xuyên')}) - Môn:"
            f" {ex_info.get('subject','')} (Lớp: {ex_info.get('target_class','ALL')})"
        ):
          st.write(
              f"⏱️ **Thời gian mở:** {ex_info.get('start_time','N/A')} ➔"
              f" {ex_info.get('end_time','N/A')} | **Thời gian làm bài:**"
              f" {ex_info.get('duration', 45)} phút | **Người tạo:**"
              f" {ex_info.get('created_by','admin')}"
          )
          st.write(
              f"⚙️ **Cấu trúc:** {ex_info.get('num_mcq',0)} câu Trắc nghiệm"
              f" ({ex_info.get('mcq_points',0)}đ) | {ex_info.get('num_tf',0)} câu"
              f" Đúng/Sai ({ex_info.get('tf_points',0)}đ) |"
              f" {ex_info.get('num_essay',0)} câu {e_lbl}"
              f" ({ex_info.get('essay_points',0)}đ)"
          )

          b_col1, b_col2, b_col3, b_col4 = st.columns(4)

          status_str = '🟢 Đang Mở' if ex_info['is_active'] else '🔴 Đang Đóng'
          if b_col1.button(f'{status_str}', key=f'toggle_{ex_id}'):
            exams_db[ex_id]['is_active'] = not exams_db[ex_id]['is_active']
            save_exams(exams_db)
            st.rerun()

          show_q_key = f'show_q_{ex_id}'
          if b_col2.button('👁️ Xem Nội Dung Đề', key=f'btn_q_{ex_id}'):
            st.session_state[show_q_key] = not st.session_state.get(
                show_q_key, False
            )

          show_st_key = f'show_st_{ex_id}'
          if b_col3.button('📊 Xem Thống Kê Điểm', key=f'btn_st_{ex_id}'):
            st.session_state[show_st_key] = not st.session_state.get(
                show_st_key, False
            )

          if b_col4.button('🗑️ Xóa Đề Thi', key=f'del_{ex_id}'):
            del exams_db[ex_id]
            save_exams(exams_db)
            st.rerun()

          if st.session_state.get(show_q_key, False):
            st.markdown('---')
            st.markdown('##### 📝 NỘI DUNG CÂU HỎI TRONG ĐỀ THI')
            bank_path = os.path.join(BANK_FOLDER, ex_info.get('bank_file', ''))
            mcq_p, tf_p, es_p = parse_question_bank(bank_path)
            st.info(f"File ngân hàng đề gốc: `{ex_info.get('bank_file')}`")

            if mcq_p and ex_info.get('num_mcq', 0) > 0:
              st.write('**I. Phần Trắc Nghiệm:**')
              for idx, q in enumerate(mcq_p[: ex_info['num_mcq']], 1):
                st.write(
                    f"**Câu {idx} [{q.get('chapter')}] [{q.get('level')}]:**"
                    f" {q['question']}"
                )
                if q.get('images'):
                  for img_p in q['images']:
                    if os.path.exists(img_p):
                      st.image(img_p, width=400)
                for opt in q.get('options', []):
                  mark = ' ✅ (Đáp án đúng)' if opt.get('is_correct') else ''
                  st.write(f"- {opt['text']}{mark}")

            if tf_p and ex_info.get('num_tf', 0) > 0:
              st.write('**II. Phần Đúng / Sai:**')
              for idx, q in enumerate(tf_p[: ex_info['num_tf']], 1):
                st.write(
                    f"**Câu {idx} [{q.get('chapter')}] [{q.get('level')}]:**"
                    f" {q['question']}"
                )
                if q.get('images'):
                  for img_p in q['images']:
                    if os.path.exists(img_p):
                      st.image(img_p, width=400)
                for sub in q.get('sub_items', []):
                  ans_str = 'ĐÚNG' if sub.get('answer') else 'SAI'
                  st.write(f"- {sub['text']} -> **[{ans_str}]**")

            if es_p and ex_info.get('num_essay', 0) > 0:
              st.write(f'**III. Phần {e_lbl}:**')
              for idx, q in enumerate(es_p[: ex_info['num_essay']], 1):
                st.write(
                    f"**Câu {idx} [{q.get('chapter')}] [{q.get('level')}]:**"
                    f" {q['question']}"
                )
                if q.get('images'):
                  for img_p in q['images']:
                    if os.path.exists(img_p):
                      st.image(img_p, width=400)
                if q.get('sample_answer'):
                  st.write(f"  *Đáp án chuẩn:* `{q['sample_answer']}`")

          if st.session_state.get(show_st_key, False):
            st.markdown('---')
            st.markdown(f'##### 📊 KẾT QUẢ & THỐNG KÊ BÀI THI [{ex_code_disp}]')
            ex_submissions = []
            for f_name in os.listdir(SUBMIT_FOLDER):
              if f_name.endswith('.json'):
                try:
                  with open(
                      os.path.join(SUBMIT_FOLDER, f_name), 'r', encoding='utf-8'
                  ) as f:
                    sdata = json.load(f)
                    real_c = get_real_exam_code(sdata, exams_db)
                    if real_c == ex_code_disp or (
                        sdata.get('exam_title') == ex_info['title']
                        and sdata.get('subject') == ex_info.get('subject')
                    ):
                      ex_submissions.append(sdata)
                except Exception:
                  pass

            if not ex_submissions:
              st.warning('Chưa có bài nộp nào cho đề thi này.')
            else:
              scores = [s.get('total_score', 0.0) for s in ex_submissions]
              avg_sc = round(sum(scores) / len(scores), 2)
              max_sc = max(scores)
              min_sc = min(scores)
              pass_r = round(
                  sum(1 for sc in scores if sc >= 5.0) / len(scores) * 100, 1
              )

              m1, m2, m3, m4 = st.columns(4)
              m1.metric('Tổng bài nộp', f'{len(scores)} bài')
              m2.metric('Điểm trung bình', f'{avg_sc}đ')
              m3.metric('Điểm Cao/Thấp nhất', f'{max_sc}đ / {min_sc}đ')
              m4.metric('Tỷ lệ Đạt (≥5.0)', f'{pass_r}%')

              tbl_data = []
              for s in ex_submissions:
                tbl_data.append({
                    'Mã HS': s.get('student_id'),
                    'Họ Tên': s.get('student_name'),
                    'Điểm Tự Động': s.get('auto_score'),
                    'Tổng Điểm': s.get('total_score'),
                    'Trạng Thái': (
                        '🟢 Đã chấm' if s.get('is_graded', True) else '⏳ Chờ chấm'
                    ),
                    'Thời Gian Nộp': s.get('timestamp'),
                })
              st.dataframe(tbl_data, use_container_width=True)

  # TAB 4: QUẢN LÝ LỚP HỌC & HỌC SINH
  with tabs[3]:
    st.subheader('🏫 QUẢN LÝ LỚP HỌC & HỌC SINH')
    all_students = [
        k for k, v in users_db.items() if v.get('role') == 'student'
    ]
    all_classes = sorted(
        list(
            set(
                users_db[k].get('class_name', '')
                for k in all_students
                if users_db[k].get('class_name')
            )
        )
    )

    if user['role'] == 'teacher' and user.get('managed_classes'):
      selectable_classes = [
          c for c in all_classes if c in user['managed_classes']
      ]
      if not selectable_classes:
        selectable_classes = user['managed_classes']
    else:
      selectable_classes = all_classes if all_classes else ['10A1']

    selected_class_view = st.selectbox(
        'Chọn Lớp Học để quản lý:', options=selectable_classes
    )

    st.markdown('---')
    exp_ex, exp_man = st.tabs([
        '📊 Nhập Học Sinh Từ File Excel',
        '➕ Thêm Thủ Công 1 Học Sinh',
    ])

    # NHẬP HỌC SINH TỪ EXCEL & TỰ ĐỘNG PHÂN LỚP
    with exp_ex:
      st.write(
          'Tải lên file Excel chứa danh sách học sinh có các cột:'
          ' **MHS/Mã/ID**, **Tên/Họ tên**, **Lớp**, **SĐT/Email**. Mật khẩu ban'
          ' đầu là `123` (Học sinh có thể tự thay đổi).'
      )
      uploaded_excel = st.file_uploader(
          'Chọn file Excel danh sách học sinh:',
          type=['xlsx', 'xls', 'csv'],
          key=f'excel_uploader_{st.session_state.excel_uploader_key}',
      )

      if uploaded_excel is not None:
        try:
          if uploaded_excel.name.endswith('.csv'):
            df = pd.read_csv(uploaded_excel)
          else:
            df = pd.read_excel(uploaded_excel)

          df.columns = [str(c).strip().lower() for c in df.columns]

          id_col = next(
              (
                  c
                  for c in df.columns
                  if 'mhs' in c or 'mã' in c or 'id' in c or 'username' in c
              ),
              None,
          )
          name_col = next(
              (
                  c
                  for c in df.columns
                  if 'tên' in c or 'name' in c or 'họ' in c
              ),
              None,
          )
          class_col = next(
              (c for c in df.columns if 'lớp' in c or 'class' in c), None
          )
          contact_col = next(
              (
                  c
                  for c in df.columns
                  if 'sđt' in c
                  or 'email' in c
                  or 'contact' in c
                  or 'liên' in c
              ),
              None,
          )

          if not id_col or not name_col:
            st.error(
                "❌ File Excel cần chứa các cột bắt buộc: 'Mã học sinh' (hoặc MHS)"
                " và 'Họ và tên'!"
            )
          else:
            count_added = 0
            for _, row in df.iterrows():
              hs_id = str(row[id_col]).strip()
              hs_name = str(row[name_col]).strip()

              # Lấy lớp từ Excel, nếu trống sẽ lấy lớp đang chọn
              hs_class = (
                  str(row[class_col]).strip()
                  if class_col and pd.notna(row[class_col])
                  else selected_class_view
              )
              hs_contact = (
                  str(row[contact_col]).strip()
                  if contact_col and pd.notna(row[contact_col])
                  else ''
              )

              if hs_id and hs_id != 'nan' and hs_name and hs_name != 'nan':
                # Giữ nguyên mật khẩu đã đổi của học sinh nếu đã có tài khoản
                existing_pass = users_db.get(hs_id, {}).get('password', '123')
                users_db[hs_id] = {
                    'name': hs_name,
                    'password': existing_pass,
                    'role': 'student',
                    'class_name': hs_class,
                    'contact': hs_contact,
                }
                count_added += 1

            save_users(users_db)
            st.session_state.excel_uploader_key += 1
            st.toast(
                f'✅ Đã nhập và phân lớp thành công {count_added} học sinh!'
            )
            st.rerun()
        except Exception as e:
          st.error(f'❌ Lỗi đọc file Excel: {e}')

    with exp_man:
      c1, c2 = st.columns(2)
      with c1:
        new_hs_id = st.text_input('Mã Học sinh (Tên đăng nhập):')
        new_hs_name = st.text_input('Họ và tên:')
      with c2:
        new_hs_pass = st.text_input('Mật khẩu:', value='123', disabled=True)
        new_hs_contact = st.text_input('Email / SĐT:')

      if st.button('➕ Thêm Học Sinh Này'):
        if new_hs_id in users_db:
          st.error('❌ Mã học sinh này đã tồn tại!')
        elif new_hs_id and new_hs_name:
          users_db[new_hs_id] = {
              'name': new_hs_name,
              'password': '123',
              'role': 'student',
              'class_name': selected_class_view,
              'contact': new_hs_contact,
          }
          save_users(users_db)
          st.toast(
              f'✅ Đã thêm học sinh {new_hs_name} vào lớp {selected_class_view}!'
          )
          st.rerun()

    st.markdown('---')
    st.subheader(f'📋 DANH SÁCH HỌC SINH LỚP {selected_class_view}')

    class_students = [
        {
            'Mã HS': k,
            'Họ Tên': v['name'],
            'Lớp': v.get('class_name', selected_class_view),
            'SĐT/Email': v.get('contact', ''),
        }
        for k, v in users_db.items()
        if v.get('role') == 'student'
        and v.get('class_name') == selected_class_view
    ]

    # SẮP XẾP TĂNG DẦN THEO MÃ HỌC SINH (MHS)
    class_students = sorted(class_students, key=lambda x: str(x['Mã HS']))

    if class_students:
      st.dataframe(class_students, use_container_width=True)
    else:
      st.info(f'Chưa có học sinh nào trong lớp {selected_class_view}.')

  # TAB 5: QUẢN LÝ GIÁO VIÊN & SAO LƯU DỮ LIỆU
  if user['role'] == 'admin' or user.get('is_head'):
    with tabs[4]:
      if user['role'] == 'admin':
        st.subheader('💾 SAO LƯU & KHÔI PHỤC DANH SÁCH TÀI KHOẢN (AN TOÀN)')

        bk_col1, bk_col2 = st.columns(2)
        with bk_col1:
          st.write('**1. Tải về File sao lưu người dùng (`users.json`):**')
          users_json_bytes = json.dumps(
              users_db, ensure_ascii=False, indent=4
          ).encode('utf-8')
          st.download_button(
              label='📥 Tải Về File Sao Lưu users.json',
              data=users_json_bytes,
              file_name='users.json',
              mime='application/json',
          )

        with bk_col2:
          st.write('**2. Khôi phục từ File sao lưu máy tính:**')
          restore_file = st.file_uploader(
              'Chọn file users.json từ máy tính:',
              type=['json'],
              key=f'restore_uploader_{st.session_state.restore_uploader_key}',
          )
          if restore_file is not None:
            try:
              imported_users = json.load(restore_file)
              save_users(imported_users)
              st.session_state.restore_uploader_key += 1
              st.toast('✅ Đã khôi phục toàn bộ danh sách người dùng!')
              st.rerun()
            except Exception as e:
              st.error(f'❌ File không hợp lệ: {e}')

        st.markdown('---')
        st.subheader('🎲 PHÂN CÔNG NGẪU NHIÊN CHẤM THI CUỐI KỲ')

        if st.button('⚡ Phân Công Ngẫu Nhiên Tất Cả Bài Thi Cuối Kỳ'):
          unassigned_count = 0
          teachers_by_subject = {}

          for u_id, u_info in users_db.items():
            if u_info.get('role') == 'teacher':
              subj = u_info.get('subject')
              teachers_by_subject.setdefault(subj, []).append(u_id)

          for f_name in os.listdir(SUBMIT_FOLDER):
            if f_name.endswith('.json'):
              f_path = os.path.join(SUBMIT_FOLDER, f_name)
              with open(f_path, 'r', encoding='utf-8') as f:
                s_data = json.load(f)

              if (
                  s_data.get('exam_type') == 'Cuối kỳ'
                  and not s_data.get('assigned_teacher')
              ):
                subj = s_data.get('subject')
                available_teachers = teachers_by_subject.get(subj, [])
                if available_teachers:
                  assigned_gv = random.choice(available_teachers)
                  s_data['assigned_teacher'] = assigned_gv
                  with open(f_path, 'w', encoding='utf-8') as f:
                    json.dump(s_data, f, ensure_ascii=False, indent=4)
                  unassigned_count += 1

          st.success(
              f'✅ Đã phân công thành công {unassigned_count} bài thi Cuối kỳ cho'
              ' các giáo viên bộ môn!'
          )

        st.markdown('---')

      st.subheader('📋 DANH SÁCH GIÁO VIÊN TRONG HỆ THỐNG')
      teacher_table = []
      for gv_id, gv_info in users_db.items():
        if gv_info.get('role') == 'teacher':
          if user['role'] == 'admin' or gv_info.get('subject') == user.get(
              'subject'
          ):
            teacher_table.append({
                'Mã GV': gv_id,
                'Họ và Tên': gv_info.get('name', ''),
                'Môn Giảng Dạy': gv_info.get('subject', ''),
                'Trưởng Bộ Môn': 'Có' if gv_info.get('is_head') else 'Không',
                'Các Lớp Quản Lý': ', '.join(
                    gv_info.get('managed_classes', [])
                ),
                'SĐT / Email': gv_info.get('contact', ''),
            })

      if teacher_table:
        st.dataframe(teacher_table, use_container_width=True)
      else:
        st.info('Chưa có dữ liệu giáo viên.')

      st.markdown('---')
      st.subheader('✏️ CẬP NHẬT CÁC LỚP QUẢN LÝ CHO GIÁO VIÊN')

      editable_teachers = {
          f"{gv_info.get('name')} ({gv_id}) - Môn: {gv_info.get('subject')}": (
              gv_id
          )
          for gv_id, gv_info in users_db.items()
          if gv_info.get('role') == 'teacher'
          and (
              user['role'] == 'admin'
              or gv_info.get('subject') == user.get('subject')
          )
      }

      if editable_teachers:
        selected_gv_label = st.selectbox(
            'Chọn Giáo viên cần chỉnh sửa lớp quản lý:',
            options=list(editable_teachers.keys()),
        )
        target_gv_id = editable_teachers[selected_gv_label]
        target_gv_info = users_db[target_gv_id]

        current_classes_str = ', '.join(
            target_gv_info.get('managed_classes', [])
        )
        new_classes_input = st.text_input(
            f"Các Lớp Quản Lý Mới cho {target_gv_info.get('name')} (phân cách bằng"
            ' dấu phẩy):',
            value=current_classes_str,
            key=f'edit_classes_in_{target_gv_id}',
        )

        if st.button('💾 Cập Nhật Lớp Quản Lý', type='primary'):
          updated_classes = [
              c.strip() for c in new_classes_input.split(',') if c.strip()
          ]
          users_db[target_gv_id]['managed_classes'] = updated_classes
          save_users(users_db)
          st.success(
              f"✅ Đã cập nhật thành công danh sách lớp quản lý cho"
              f" {target_gv_info.get('name')}!"
          )
          st.rerun()

      if user['role'] == 'admin':
        st.markdown('---')
        st.subheader('➕ TẠO TÀI KHOẢN GIÁO VIÊN MỚI')

        cg1, cg2 = st.columns(2)
        with cg1:
          new_gv_id = st.text_input('Mã Giáo viên:')
          new_gv_name = st.text_input('Họ tên Giáo viên:')
          new_gv_subject = st.selectbox(
              'Môn học giảng dạy:', list(SUBJECT_PREFIXES.keys())
          )
        with cg2:
          new_gv_pass = st.text_input('Mật khẩu:', value='123', key='gv_pass_in')
          new_gv_is_head = st.checkbox(
              'Trưởng bộ môn (Có quyền tạo đề Cuối kỳ & Quản lý bộ môn)'
          )
          new_gv_classes_str = st.text_input(
              'Các Lớp Quản Lý:', value='10A1, 10A2'
          )

        if st.button('➕ Tạo Giáo Viên'):
          if new_gv_id in users_db:
            st.error('❌ Mã giáo viên này đã tồn tại!')
          elif new_gv_id and new_gv_name:
            classes_list = [
                c.strip() for c in new_gv_classes_str.split(',') if c.strip()
            ]
            users_db[new_gv_id] = {
                'name': new_gv_name,
                'password': new_gv_pass,
                'role': 'teacher',
                'is_head': new_gv_is_head,
                'subject': new_gv_subject,
                'managed_classes': classes_list,
                'contact': '',
            }
            save_users(users_db)
            st.success(f'✅ Tạo thành công Giáo viên {new_gv_name}!')
            st.rerun()


# ==================== GIAO DIỆN HỌC SINH ====================
elif user['role'] == 'student':

  # SIDEBAR HỌC SINH (BẢNG ĐIỀU HƯỚNG TỰ ĐỘNG ĐỔI MÀU VÀNG REALTIME)
  with st.sidebar:
    if st.session_state.get('exam_started', False):
      exam = st.session_state.exam_data
      exam_info = st.session_state.current_exam_info
      e_label = (
          'Trả Lời Ngắn'
          if exam_info.get('subject') in SHORT_ANSWER_SUBJECTS
          else 'Tự Luận'
      )

      start_time_ts = st.session_state.get('start_time', time.time())
      duration_seconds = exam_info['duration'] * 60
      end_by_duration_ms = int((start_time_ts + duration_seconds) * 1000)

      end_str = exam_info.get('end_time')
      if end_str:
        try:
          end_dt = datetime.datetime.strptime(end_str, '%Y-%m-%d %H:%M')
          end_by_schedule_ms = int(end_dt.timestamp() * 1000)
          final_end_ms = min(end_by_duration_ms, end_by_schedule_ms)
        except Exception:
          final_end_ms = end_by_duration_ms
      else:
        final_end_ms = end_by_duration_ms

      timer_code = f"""
            <div style="font-family: Arial; text-align: center; background-color: #fff3cd; border: 2px solid #ffeba2; padding: 10px; border-radius: 8px; margin-bottom: 12px;">
                <div style="font-size: 13px; color: #856404; font-weight: bold;">⏱️ THỜI GIAN CÒN LẠI</div>
                <div id="timer" style="font-size: 24px; font-weight: bold; color: #d9534f;">--:--</div>
            </div>
            <script>
            var endTime = {final_end_ms};
            var timerInterval = setInterval(function() {{
                var now = new Date().getTime();
                var distance = endTime - now;
                if (distance <= 0) {{
                    clearInterval(timerInterval);
                    document.getElementById("timer").innerHTML = "HẾT GIỜ!";
                    var btns = window.parent.document.querySelectorAll('button');
                    for (var i = 0; i < btns.length; i++) {{
                        if (btns[i].innerText.includes('Nộp Bài')) {{
                            btns[i].click();
                            break;
                        }}
                    }}
                }} else {{
                    var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
                    var seconds = Math.floor((distance % (1000 * 60)) / 1000);
                    document.getElementById("timer").innerHTML = 
                        (minutes < 10 ? "0" : "") + minutes + ":" + (seconds < 10 ? "0" : "") + seconds;
                }}
            }}, 1000);
            </script>
            """
      components.html(timer_code, height=85)

      st.markdown('---')
      st.subheader('📌 BẢNG ĐIỀU HƯỚNG')

      if exam['mcq']:
        st.write('**I. Trắc Nghiệm**')
        nav_html = "<div style='display: flex; flex-wrap: wrap; gap: 6px;'>"
        for idx in range(1, len(exam['mcq']) + 1):
          val = st.session_state.get(f'mcq_{idx}')
          is_ans = val is not None
          bg_color = '#ffc107' if is_ans else '#6c757d'
          text_color = 'black' if is_ans else 'white'
          nav_html += f"""<a href="#q_mcq_{idx}" style="display:inline-block; width:32px; height:32px; line-height:32px; text-align:center; background-color:{bg_color}; color:{text_color}; border-radius:6px; font-weight:bold; text-decoration:none;">{idx}</a>"""
        nav_html += '</div>'
        st.markdown(nav_html, unsafe_allow_html=True)

      if exam['tf']:
        st.write('**II. Đúng Sai**')
        nav_html_tf = "<div style='display: flex; flex-wrap: wrap; gap: 6px;'>"
        for idx, q in enumerate(exam['tf'], 1):
          num_subs = len(q.get('sub_items', []))
          is_ans = any(
              st.session_state.get(f'tf_{idx}_{sub}') is not None
              for sub in range(num_subs)
          )
          bg_color = '#ffc107' if is_ans else '#6c757d'
          text_color = 'black' if is_ans else 'white'
          nav_html_tf += f"""<a href="#q_tf_{idx}" style="display:inline-block; width:32px; height:32px; line-height:32px; text-align:center; background-color:{bg_color}; color:{text_color}; border-radius:6px; font-weight:bold; text-decoration:none;">{idx}</a>"""
        nav_html_tf += '</div>'
        st.markdown(nav_html_tf, unsafe_allow_html=True)

      if exam['essay']:
        st.write(f'**III. {e_label}**')
        nav_html_es = "<div style='display: flex; flex-wrap: wrap; gap: 6px;'>"
        for idx in range(1, len(exam['essay']) + 1):
          val = st.session_state.get(f'essay_{idx}', '')
          is_ans = bool(str(val).strip())
          bg_color = '#ffc107' if is_ans else '#6c757d'
          text_color = 'black' if is_ans else 'white'
          nav_html_es += f"""<a href="#q_essay_{idx}" style="display:inline-block; width:34px; height:34px; line-height:34px; text-align:center; background-color:{bg_color}; color:{text_color}; border-radius:6px; font-weight:bold; text-decoration:none;">{idx}</a>"""
        nav_html_es += '</div>'
        st.markdown(nav_html_es, unsafe_allow_html=True)

    else:
      st.markdown('### 📚 DANH SÁCH MÔN HỌC')
      student_subjects = ['Tất cả môn'] + list(SUBJECT_PREFIXES.keys())
      selected_subject = st.radio('Chọn môn học:', student_subjects)

      st.markdown('---')
      st.markdown('### 📊 TỔNG HỢP KẾT QUẢ THEO MÔN')

      student_class = user.get('class_name', '')
      summary_list = []

      for subj in SUBJECT_PREFIXES.keys():
        subj_exams = [
            v
            for v in exams_db.values()
            if v.get('is_active', True)
            and v.get('subject') == subj
            and (
                v.get('target_class') == 'ALL'
                or v.get('target_class') == student_class
            )
        ]
        tot = len(subj_exams)
        if tot == 0:
          continue

        done_cnt = 0
        score_sum = 0.0
        for ex in subj_exams:
          subm = get_student_exam_submission(user['username'], ex.get('code'))
          if subm:
            done_cnt += 1
            score_sum += float(subm.get('total_score', 0.0))

        comp_pct = f'{int(round((done_cnt / tot) * 100))}%'
        avg_sc = f'{round(score_sum / done_cnt, 2):.2f}' if done_cnt > 0 else '-'

        summary_list.append(
            {'Môn': subj, 'Điểm TB': avg_sc, '% Hoàn thành': comp_pct}
        )

      if summary_list:
        st.dataframe(summary_list, use_container_width=True, hide_index=True)

  # KẾT QUẢ SAU KHI NỘP BÀI
  if st.session_state.submitted_data:
    res = st.session_state.submitted_data
    st.balloons()
    st.title('🎉 BÀI THI ĐÃ ĐƯỢC NỘP THÀNH CÔNG')

    st.markdown(
        f"""
        <div style="background-color: #f8f9fa; border: 2px solid #28a745; padding: 20px; border-radius: 10px; margin-top: 15px;">
            <h2 style="color: #28a745; margin-top:0;">📋 KẾT QUẢ BÀI THI: {res['exam_title']} ({res.get('exam_type','Giữa kỳ')})</h2>
            <hr>
            <p style="font-size: 18px;"><b>Môn học:</b> {res['subject']} | <b>Mã bài thi:</b> {res['exam_code']}</p>
            <p style="font-size: 18px;"><b>Thời gian nộp bài:</b> {res['timestamp']}</p>
            <p style="font-size: 18px;"><b>Số câu làm đúng:</b> <span style="color: #28a745; font-weight: bold;">{res['correct_count']} câu</span></p>
            <p style="font-size: 22px;"><b>Tổng điểm đạt được:</b> <span style="color: #d9534f; font-weight: bold;">{res['total_score']} Điểm</span></p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    if st.button('🔄 Quay Lại Danh Sách Bài Thi'):
      clear_exam_session_state()
      st.session_state.submitted_data = None
      st.session_state.exam_started = False
      st.session_state.confirm_exam_id = None
      st.rerun()

    st.stop()

  # CỬA SỔ XÁC NHẬN VÀO THI & TỰ ĐỘNG KHÔI PHỤC BÀI DỞ DÙ OUT RA VÀO LẠI
  if not st.session_state.get('exam_started', False) and st.session_state.get(
      'confirm_exam_id'
  ):
    ex_id = st.session_state.confirm_exam_id
    exam_info = exams_db.get(ex_id)

    if exam_info:
      draft = load_student_draft(user['username'], exam_info['code'])

      st.markdown('<br><br>', unsafe_allow_html=True)
      col_m1, col_m2, col_m3 = st.columns([1, 2.2, 1])

      with col_m2:
        st.markdown("<div class='confirm-modal'>", unsafe_allow_html=True)

        if draft:
          st.subheader('📌 PHÁT HIỆN BÀI LÀM CHƯA HOÀN THÀNH')
          st.warning(
              f"Hệ thống tìm thấy bài làm dở lưu lúc **{draft.get('updated_at')}**."
              " Toàn bộ đáp án bạn đã làm sẽ được giữ nguyên 100%!"
          )
          st.write(
              f"Bài thi: **[{exam_info.get('code')}] {exam_info['title']}**"
          )
          st.write(f"- **Môn:** {exam_info.get('subject')}")
          st.write(f"- **Thời gian làm bài:** {exam_info['duration']} phút")

          st.markdown('<br>', unsafe_allow_html=True)
          col_b1, col_b2, col_b3 = st.columns(3)
          with col_b1:
            if st.button(
                '▶️ Làm Tiếp Bài Dở', use_container_width=True, type='primary'
            ):
              restore_exam_from_draft(draft)
              st.rerun()

          with col_b2:
            if st.button('🔄 Làm Bài Mới', use_container_width=True):
              delete_student_draft(user['username'], exam_info['code'])
              clear_exam_session_state()

              mcq_p, tf_p, es_p = parse_question_bank(
                  os.path.join(BANK_FOLDER, exam_info['bank_file'])
              )
              random.seed(int(time.time() * 1000) + hash(user['username']))

              selected_mcq = sample_questions_balanced(
                  mcq_p, exam_info['num_mcq']
              )
              selected_tf = sample_questions_balanced(
                  tf_p, exam_info['num_tf']
              )
              selected_essay = sample_questions_balanced(
                  es_p, exam_info['num_essay']
              )

              mcq_opts_map = {}
              for q in selected_mcq:
                fmt_opts, corr_lbl = process_mcq_question(q)
                st.session_state[f'mcq_opts_{q["id"]}'] = fmt_opts
                st.session_state[f'mcq_corr_{q["id"]}'] = corr_lbl
                mcq_opts_map[str(q['id'])] = {
                    'fmt_opts': fmt_opts,
                    'corr_lbl': corr_lbl,
                }

              st.session_state.exam_data = {
                  'mcq': selected_mcq,
                  'tf': selected_tf,
                  'essay': selected_essay,
              }
              st.session_state.start_time = time.time()
              st.session_state.exam_started = True
              st.session_state.current_exam_info = exam_info
              st.session_state.confirm_exam_id = None
              st.rerun()

          with col_b3:
            if st.button('❌ Hủy', use_container_width=True):
              st.session_state.confirm_exam_id = None
              st.rerun()

        else:
          st.subheader('⚠️ XÁC NHẬN VÀO THI')
          st.write(
              f"Bạn có chắc chắn muốn bắt đầu làm bài thi:"
              f" **[{exam_info.get('code')}] {exam_info['title']}**?"
          )
          st.write(f"- **Môn:** {exam_info.get('subject')}")
          st.write(f"- **Thời gian làm bài:** {exam_info['duration']} phút")

          st.markdown('<br>', unsafe_allow_html=True)
          col_b1, col_b2 = st.columns(2)
          with col_b1:
            if st.button(
                '🚀 Bắt Đầu Vào Thi', use_container_width=True, type='primary'
            ):
              clear_exam_session_state()

              mcq_p, tf_p, es_p = parse_question_bank(
                  os.path.join(BANK_FOLDER, exam_info['bank_file'])
              )
              random.seed(int(time.time() * 1000) + hash(user['username']))

              selected_mcq = sample_questions_balanced(
                  mcq_p, exam_info['num_mcq']
              )
              selected_tf = sample_questions_balanced(
                  tf_p, exam_info['num_tf']
              )
              selected_essay = sample_questions_balanced(
                  es_p, exam_info['num_essay']
              )

              mcq_opts_map = {}
              for q in selected_mcq:
                fmt_opts, corr_lbl = process_mcq_question(q)
                st.session_state[f'mcq_opts_{q["id"]}'] = fmt_opts
                st.session_state[f'mcq_corr_{q["id"]}'] = corr_lbl
                mcq_opts_map[str(q['id'])] = {
                    'fmt_opts': fmt_opts,
                    'corr_lbl': corr_lbl,
                }

              st.session_state.exam_data = {
                  'mcq': selected_mcq,
                  'tf': selected_tf,
                  'essay': selected_essay,
              }
              st.session_state.start_time = time.time()
              st.session_state.exam_started = True
              st.session_state.current_exam_info = exam_info
              st.session_state.confirm_exam_id = None
              st.rerun()

          with col_b2:
            if st.button('❌ Thoát', use_container_width=True):
              st.session_state.confirm_exam_id = None
              st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)

    st.stop()

  # DANH SÁCH ĐỀ THI BÊN HỌC SINH
  if not st.session_state.get('exam_started', False):
    st.title(
        f"📝 DANH SÁCH ĐỀ THI - {selected_subject.upper() if selected_subject != 'Tất cả môn' else 'TẤT CẢ MÔN'}"
    )
    student_class = user.get('class_name', '')

    visible_exams = {}
    for k, v in exams_db.items():
      if v.get('is_active', True) and (
          v.get('target_class') == 'ALL'
          or v.get('target_class') == student_class
      ):
        if (
            selected_subject == 'Tất cả môn'
            or v.get('subject') == selected_subject
        ):
          visible_exams[k] = v

    if not visible_exams:
      st.warning(
          f'⚠️ Hiện tại chưa có bài thi nào cho môn **{selected_subject}** dành'
          f' cho lớp {student_class}!'
      )
      st.stop()

    search_kw = st.text_input(
        '🔍 Tìm kiếm bài thi theo Mã bài thi hoặc Tên bài thi:', value=''
    ).strip()

    filtered_exams = {}
    for k, v in visible_exams.items():
      kw_lower = search_kw.lower()
      if (
          kw_lower in v.get('code', '').lower()
          or kw_lower in v.get('title', '').lower()
      ):
        filtered_exams[k] = v

    if not filtered_exams:
      st.info('Không tìm thấy bài thi nào phù hợp.')
      st.stop()

    st.markdown('---')
    now_dt = datetime.datetime.now()

    def exam_sort_key(ex_id):
      ex = filtered_exams[ex_id]
      subm = get_student_exam_submission(user['username'], ex.get('code'))
      has_done = subm is not None

      end_str = ex.get('end_time')
      end_dt = (
          parse_timestamp(end_str) if end_str else datetime.datetime.max
      )
      is_expired = now_dt > end_dt if end_str else False

      if not has_done and not is_expired:
        return (0, end_dt, ex.get('code', ''))
      elif has_done:
        return (1, end_dt, ex.get('code', ''))
      else:
        return (2, end_dt, ex.get('code', ''))

    sorted_exam_keys = sorted(filtered_exams.keys(), key=exam_sort_key)

    for ex_id in sorted_exam_keys:
      ex_info = filtered_exams[ex_id]
      submission = get_student_exam_submission(
          user['username'], ex_info['code']
      )
      has_done = submission is not None
      draft = load_student_draft(user['username'], ex_info['code'])

      start_str = ex_info.get('start_time')
      end_str = ex_info.get('end_time')

      is_not_opened_yet = False
      is_expired = False

      if start_str and end_str:
        try:
          start_dt = parse_timestamp(start_str)
          end_dt = parse_timestamp(end_str)
          if now_dt < start_dt:
            is_not_opened_yet = True
          elif now_dt > end_dt:
            is_expired = True
        except Exception:
          pass

      col_e1, col_e2, col_e3 = st.columns([3.5, 1.2, 1.3])
      with col_e1:
        st.write(
            f"**[{ex_info.get('code','DE')}] {ex_info['title']}** - Môn:"
            f" {ex_info.get('subject','')} (Lớp: {ex_info.get('target_class','ALL')})"
        )
        st.caption(f'⏱️ Hạn mở đề: {start_str} ➔ {end_str}')
      with col_e2:
        if has_done:
          st.markdown(
              "<span style='background-color:#d4edda; color:#155724;"
              " padding:4px 8px; border-radius:4px;'>🟢 Đã Làm</span>",
              unsafe_allow_html=True,
          )
        elif draft:
          st.markdown(
              "<span style='background-color:#fff3cd; color:#856404;"
              " padding:4px 8px; border-radius:4px;'>🟡 Đang Làm Dở</span>",
              unsafe_allow_html=True,
          )
        elif is_not_opened_yet:
          st.markdown(
              "<span style='background-color:#fff3cd; color:#856404;"
              " padding:4px 8px; border-radius:4px;'>🔒 Chưa Mở</span>",
              unsafe_allow_html=True,
          )
        elif is_expired:
          st.markdown(
              "<span style='background-color:#f8d7da; color:#721c24;"
              " padding:4px 8px; border-radius:4px;'>⚪ Chưa Làm (Quá"
              ' Hạn)</span>',
              unsafe_allow_html=True,
          )
        else:
          st.markdown(
              "<span style='background-color:#e2e3e5; color:#383d41;"
              " padding:4px 8px; border-radius:4px;'>⚪ Chưa Làm</span>",
              unsafe_allow_html=True,
          )
      with col_e3:
        if has_done:
          if ex_info.get('exam_type') == 'Cuối kỳ' and not submission.get(
              'is_graded'
          ):
            st.markdown(
                "<span style='color:#fd7e14; font-weight:bold;'>⏳ Chờ công"
                ' bố</span>',
                unsafe_allow_html=True,
            )
          else:
            tot_sc = submission.get('total_score', 0.0)
            st.markdown(
                f"<span style='background-color:#cce5ff; color:#004085;"
                f" padding:4px 8px; border-radius:4px;'>🏆 {tot_sc}đ</span>",
                unsafe_allow_html=True,
            )
        elif not is_not_opened_yet and not is_expired:
          btn_label = '▶️ Làm Tiếp' if draft else '▶️ Vào Thi'
          if st.button(btn_label, key=f'btn_start_direct_{ex_id}'):
            if draft:
              restore_exam_from_draft(draft)
              st.rerun()
            else:
              st.session_state.confirm_exam_id = ex_id
              st.rerun()

      st.markdown('---')

    st.stop()

  # MÀN HÌNH ĐANG LÀM BÀI THI CỦA HỌC SINH (LƯU & PHỦ ĐÁP ÁN CŨ KHÔNG MẤT DỮ LIỆU)
  exam = st.session_state.exam_data
  exam_info = st.session_state.current_exam_info
  is_short_ans_subj = exam_info.get('subject') in SHORT_ANSWER_SUBJECTS

  st.subheader(
      f"📝 [{exam_info.get('code','DE')}] {exam_info['title']} - Môn:"
      f" {exam_info.get('subject')}"
  )

  user_mcq = {}
  mcq_opts_map = {}

  if exam['mcq']:
    st.header('I. Trắc Nghiệm (4 lựa chọn)')
    for idx, q in enumerate(exam['mcq'], 1):
      st.markdown(f'<div id="q_mcq_{idx}"></div>', unsafe_allow_html=True)
      st.subheader(f'Câu {idx}: {q["question"]}')

      if q.get('images'):
        for img_p in q['images']:
          if os.path.exists(img_p):
            st.image(img_p, width=450)

      opts = st.session_state.get(f'mcq_opts_{q["id"]}', [])
      corr_l = st.session_state.get(f'mcq_corr_{q["id"]}', 'A')
      mcq_opts_map[str(q['id'])] = {'fmt_opts': opts, 'corr_lbl': corr_l}

      key_mcq = f'mcq_{idx}'
      saved_val = st.session_state.get(key_mcq)

      if saved_val in opts:
        user_mcq[idx] = st.radio('Chọn đáp án:', opts, key=key_mcq)
      else:
        user_mcq[idx] = st.radio('Chọn đáp án:', opts, key=key_mcq, index=None)

  user_tf = {}
  if exam['tf']:
    st.header('II. Đúng / Sai')
    for idx, q in enumerate(exam['tf'], 1):
      st.markdown(f'<div id="q_tf_{idx}"></div>', unsafe_allow_html=True)
      st.subheader(f'Câu {idx}: {q["question"]}')

      if q.get('images'):
        for img_p in q['images']:
          if os.path.exists(img_p):
            st.image(img_p, width=450)

      user_tf[idx] = []
      for sub_idx, sub_item in enumerate(q.get('sub_items', [])):
        key_tf = f'tf_{idx}_{sub_idx}'
        saved_tf_val = st.session_state.get(key_tf)

        if saved_tf_val in ['Đúng', 'Sai']:
          ans = st.radio(
              f'{sub_item["text"]}',
              ['Đúng', 'Sai'],
              key=key_tf,
              horizontal=True,
          )
        else:
          ans = st.radio(
              f'{sub_item["text"]}',
              ['Đúng', 'Sai'],
              key=key_tf,
              index=None,
              horizontal=True,
          )

        user_tf[idx].append(ans)

  user_essay = {}
  if exam['essay']:
    section_title = 'III. Trả Lời Ngắn' if is_short_ans_subj else 'III. Tự Luận'
    st.header(section_title)
    for idx, q in enumerate(exam['essay'], 1):
      st.markdown(f'<div id="q_essay_{idx}"></div>', unsafe_allow_html=True)
      st.subheader(f'Câu {idx}: {q["question"]}')

      if q.get('images'):
        for img_p in q['images']:
          if os.path.exists(img_p):
            st.image(img_p, width=450)

      key_essay = f'essay_{idx}'
      if is_short_ans_subj:
        user_essay[idx] = st.text_input(
            'Nhập đáp án (VD: 4.5 hoặc 4,5):', key=key_essay
        )
      else:
        user_essay[idx] = st.text_area('Bài làm:', key=key_essay, height=120)

  save_student_draft(
      user['username'],
      exam_info.get('code', 'DE'),
      exam_info,
      st.session_state.get('start_time', time.time()),
      exam['mcq'],
      exam['tf'],
      exam['essay'],
      user_mcq,
      user_tf,
      user_essay,
      mcq_opts_map=mcq_opts_map,
  )

  st.markdown('---')
  submit_btn = st.button('📥 Nộp Bài Cho Giáo Viên', type='primary')

  # XỬ LÝ NỘP BÀI THẬT
  if submit_btn:
    delete_student_draft(user['username'], exam_info.get('code', 'DE'))

    mcq_score = 0.0
    correct_cnt = 0

    mcq_pts = float(exam_info.get('mcq_points', 0.0))
    mcq_unit = mcq_pts / len(exam['mcq']) if exam['mcq'] else 0.0

    mcq_details = []
    if exam['mcq']:
      for idx, q in enumerate(exam['mcq'], 1):
        u_ans = user_mcq.get(idx)
        u_label = u_ans[0] if u_ans else ''
        corr_label = st.session_state.get(f'mcq_corr_{q["id"]}', '')
        is_corr = u_label == corr_label
        if is_corr:
          mcq_score += mcq_unit
          correct_cnt += 1
        mcq_details.append({
            'question': q['question'],
            'user_ans': u_label,
            'correct': corr_label,
            'is_correct': is_corr,
        })

    tf_score = 0.0
    tf_details = []
    if exam['tf']:
      for idx, q in enumerate(exam['tf'], 1):
        sub_items = q.get('sub_items', [])
        corr_tf = [sub['answer'] for sub in sub_items]
        u_tf = [ans == 'Đúng' for ans in user_tf.get(idx, [])]
        sc, c_cnt = grade_tf(u_tf, corr_tf)
        tf_score += sc
        if c_cnt == len(sub_items) and len(sub_items) > 0:
          correct_cnt += 1
        tf_details.append({
            'question': q['question'],
            'user_tf': user_tf.get(idx, []),
            'corr_tf': corr_tf,
            'score': sc,
            'correct_count': c_cnt,
        })

    essay_details = []
    essay_score = 0.0
    essay_pts = float(exam_info.get('essay_points', 0.0))

    if exam['essay']:
      if is_short_ans_subj:
        if exam_info.get('subject') == 'Toán':
          essay_unit = 0.5
        elif exam_info.get('subject') in ['Vật Lý', 'Hóa Học']:
          essay_unit = 0.25
        else:
          essay_unit = essay_pts / len(exam['essay'])
      else:
        essay_unit = (
            essay_pts / len(exam['essay']) if len(exam['essay']) > 0 else 0.0
        )

      for idx, q in enumerate(exam['essay'], 1):
        u_txt = str(user_essay.get(idx, '')).strip()
        sample_ans = q.get('sample_answer', '').strip()

        if is_short_ans_subj:
          is_correct = check_short_answer(u_txt, sample_ans)
          q_score = essay_unit if is_correct else 0.0
          if is_correct:
            correct_cnt += 1
            essay_score += q_score
          essay_details.append({
              'question': q['question'],
              'user_text': u_txt,
              'sample_answer': sample_ans,
              'final_score': q_score,
              'is_correct': is_correct,
          })
        else:
          essay_details.append({
              'question': q['question'],
              'user_text': u_txt,
              'sample_answer': sample_ans,
              'final_score': 0.0,
          })

    if is_short_ans_subj:
      auto_score = round(mcq_score + tf_score + essay_score, 2)
      total_score = auto_score
      is_fully_graded = True
    else:
      auto_score = round(mcq_score + tf_score, 2)
      total_score = auto_score
      is_fully_graded = (
          False if float(exam_info.get('essay_points', 0)) > 0 else True
      )

    timestamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    submission_data = {
        'student_name': user['name'],
        'student_id': user['username'],
        'exam_title': exam_info['title'],
        'exam_code': exam_info.get('code', 'DE'),
        'exam_type': exam_info.get('exam_type', 'Thường xuyên'),
        'subject': exam_info.get('subject', ''),
        'created_by': exam_info.get('created_by', 'admin'),
        'assigned_teacher': (
            None
            if exam_info.get('exam_type') == 'Cuối kỳ'
            else exam_info.get('created_by')
        ),
        'timestamp': timestamp,
        'auto_score': auto_score,
        'total_score': total_score,
        'correct_count': correct_cnt,
        'essay_points': essay_pts,
        'mcq_details': mcq_details,
        'tf_details': tf_details,
        'essay_details': essay_details,
        'teacher_note': '',
        'is_graded': is_fully_graded,
    }

    file_name = f'{user["username"]}_{int(time.time())}.json'
    save_path = os.path.join(SUBMIT_FOLDER, file_name)
    with open(save_path, 'w', encoding='utf-8') as f:
      json.dump(submission_data, f, ensure_ascii=False, indent=4)

    st.session_state.submitted_data = submission_data
    st.session_state.exam_started = False
    st.rerun()
