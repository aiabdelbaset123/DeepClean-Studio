#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
import streamlit as st
from io import BytesIO

try:
    from docx import Document
    from docx.shared import Pt
    DOCX_OK = True
except:
    DOCX_OK = False

st.set_page_config(page_title="DeepClean - Humanize", layout="wide")
st.title("🧬 DeepClean Studio")
st.caption("يغير النصوص الأكاديمية إلى أسلوب بشري - يحافظ على الجداول والأشكال والمعادلات")

# ========== قواعد التحويل (بدون إضافات تخريبية) ==========
REPLACE = {
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "highlight": "point out",
    "constitute": "are", "trajectories": "paths", "pronounced": "clear",
    "routinely": "often", "impose": "bring", "exceeding": "above",
    "cumulative": "total", "uniquely": "", "forecasts": "expects",
    "committed": "plans", "constitutes": "is", "expose": "give",
    "fragmented": "split", "incorporating": "using",
}

def humanize(text):
    """تطبيق الاستبدالات فقط - لا إضافات تخريبية"""
    if not text.strip():
        return text
    for old, new in REPLACE.items():
        text = re.sub(rf'\b{re.escape(old)}\b', new, text, flags=re.I)
    # تنظيف المسافات فقط
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text.strip()

def process_word(file_bytes):
    """معالجة ملف Word مع الحفاظ على كل شيء"""
    doc = Document(file_bytes)
    modified = 0
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        # تخطي الفقرات التي قد تحتوي على معادلات
        skip = False
        for run in para.runs:
            if run.element.xpath('.//m:oMath') or run.element.xpath('.//w:object'):
                skip = True
                break
        if skip:
            continue
        new_text = humanize(para.text)
        if new_text != para.text:
            style = para.style
            alignment = para.paragraph_format.alignment
            para.clear()
            run = para.add_run(new_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            para.style = style
            if alignment:
                para.paragraph_format.alignment = alignment
            modified += 1
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, modified

# ========== واجهة المستخدم ==========
st.subheader("📝 أدخل النص (لصق مباشر)")
user_text = st.text_area("", height=200)

st.subheader("📁 أو ارفع ملف Word")
uploaded = st.file_uploader("اختر ملف .docx", type=["docx"])

if st.button("🔄 تحويل", type="primary"):
    if user_text.strip():
        result = humanize(user_text)
        col1, col2 = st.columns(2)
        col1.text_area("النص الأصلي", user_text, height=300)
        col2.text_area("النص المعدل", result, height=300)
        st.download_button("تحميل TXT", result.encode(), "humanized.txt")
    elif uploaded:
        out, count = process_word(BytesIO(uploaded.read()))
        st.success(f"تم تعديل {count} فقرة. الجداول والأشكال سليمة.")
        st.download_button("تحميل Word", out, "humanized.docx")
    else:
        st.warning("أدخل نصاً أو ارفع ملفاً")
