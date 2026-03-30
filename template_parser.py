"""
template_parser.py — 学校格式规范 AI 解析器
调用 SiliconFlow Qwen2.5-VL 从 PDF/图片中提取格式规则
"""
import os
import json
import base64
import requests
import streamlit as st
from io import BytesIO

try:
    from pdf2image import convert_from_bytes
except ImportError:
    convert_from_bytes = None

SILICONFLOW_API_KEY = st.secrets.get('SILICONFLOW_API_KEY', '') if hasattr(st, 'secrets') else os.environ.get('SILICONFLOW_API_KEY', '')
if not SILICONFLOW_API_KEY:
    SILICONFLOW_API_KEY = os.environ.get('SILICONFLOW_API_KEY', '')

SILICONFLOW_URL = "https://api.siliconflow.cn/v1/chat/completions"
MODEL = "Qwen/Qwen2.5-VL-72B-Instruct"

EXTRACT_PROMPT = """你是论文格式规范解析专家。请从图片中的学校论文格式要求中，提取所有格式规则。
严格按以下 JSON 格式输出，只输出 JSON，不要任何解释。未在文档中提及的字段输出 null。
{
  "school_name": "学校名称",
  "degree": "硕士/博士",
  "page": {
    "size": "A4 或其他",
    "margin_top_cm": 数字,
    "margin_bottom_cm": 数字,
    "margin_left_cm": 数字,
    "margin_right_cm": 数字
  },
  "body": {
    "cn_font": "中文字体名",
    "en_font": "英文字体名",
    "font_size": "字号名称如小四",
    "line_spacing": 数字如1.5,
    "first_indent_char": 数字如2
  },
  "headings": {
    "h1": {"font": "字体", "size": "字号", "bold": true/false},
    "h2": {"font": "字体", "size": "字号", "bold": true/false},
    "h3": {"font": "字体", "size": "字号", "bold": true/false},
    "numbering": "numeric 或 chapter"
  },
  "abstract": {
    "title_cn_font": "字体",
    "title_cn_size": "字号",
    "title_en_font": "字体",
    "title_en_size": "字号",
    "word_count_min": 数字,
    "word_count_max": 数字,
    "keyword_count_min": 数字,
    "keyword_count_max": 数字
  },
  "toc": {
    "title_font": "字体",
    "title_size": "字号"
  },
  "caption": {
    "font": "字体",
    "size": "字号",
    "bilingual": true/false
  },
  "header": {
    "odd_page": "奇数页页眉内容",
    "even_page": "偶数页页眉内容或auto",
    "odd_even_different": true/false
  },
  "references": {
    "min_count": 数字,
    "foreign_ratio": 小数如0.33,
    "cn_before_en": true/false
  },
  "cover_fields": ["字段1", "字段2"]
}"""


def _image_to_base64(image_bytes):
    return base64.b64encode(image_bytes).decode('utf-8')


def _call_vl_model(image_b64_list):
    """调用 Qwen2.5-VL，传入图片列表"""
    content = [{"type": "text", "text": EXTRACT_PROMPT}]
    for b64 in image_b64_list:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64}"}
        })

    resp = requests.post(
        SILICONFLOW_URL,
        headers={
            "Authorization": f"Bearer {SILICONFLOW_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": MODEL,
            "messages": [{"role": "user", "content": content}],
            "max_tokens": 4096,
        },
        timeout=60,
    )
    resp.raise_for_status()
    text = resp.json()['choices'][0]['message']['content']

    # 提取 JSON（可能包裹在 ```json ``` 中）
    if '```json' in text:
        text = text.split('```json')[1].split('```')[0]
    elif '```' in text:
        text = text.split('```')[1].split('```')[0]
    return json.loads(text.strip())


def parse_template(file_bytes, filename):
    """解析上传的格式规范文件，返回规则 dict
    支持: PDF, PNG, JPG, JPEG
    """
    ext = filename.lower().rsplit('.', 1)[-1]

    if ext == 'pdf':
        if convert_from_bytes is None:
            raise ImportError("pdf2image 未安装，无法解析 PDF 文件")
        # PDF 转图片（取前5页）
        images = convert_from_bytes(file_bytes, dpi=150, last_page=5)
        b64_list = []
        for img in images:
            buf = BytesIO()
            img.save(buf, format='PNG')
            b64_list.append(_image_to_base64(buf.getvalue()))
    elif ext in ('png', 'jpg', 'jpeg'):
        b64_list = [_image_to_base64(file_bytes)]
    else:
        raise ValueError(f"不支持的文件格式: {ext}，请上传 PDF 或图片")

    return _call_vl_model(b64_list)
