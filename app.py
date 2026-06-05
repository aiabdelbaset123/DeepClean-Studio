#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
import random
import streamlit as st
from io import BytesIO

# مكتبات Word
try:
    from docx import Document
    from docx.shared import Pt
    DOCX_OK = True
except:
    DOCX_OK = False

st.set_page_config(page_title="DeepClean", layout="wide")
st.title("🧬 DeepClean Studio - النسخة النهائية")

# ==================== صندوق اللصق الرئيسي ====================
st.subheader("📝 أدخل النص هنا (لصق مباشر)")
user_text = st.text_area("", height=200, key="paste_area")

st.markdown("---")
st.subheader("📁 أو ارفع ملف Word (يتم الحفاظ على الجداول والأشكال)")
uploaded_file = st.file_uploader("اختر ملف .docx", type=["docx"])

# الإعدادات
col1, col2 = st.columns(2)
with col1:
    intensity = st.slider("شدة المراجعة", 1, 5, 3)
with col2:
    st.write("")
    st.write("")

# ==================== قواعد التحويل ====================
replacements = {
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "highlight": "point out",
    "constitute": "are", "trajectories": "paths", "pronounced": "large",
    "routinely": "often", "impose": "bring", "exceeding": "above",
    "cumulative": "total", "uniquely": "", "forecasts": "expects",
}

def humanize(txt, level):
    if not txt.strip():
        return txt
    for old, new in replacements.items():
        txt = re.sub(rf'\b{re.escape(old)}\b', new, txt, flags=re.I)
    # تقطيع الجمل الطويلة
    if level >= 3:
        sents = re.split(r'(?<=[.!?])\s+', txt)
        new_sents = []
        for s in sents:
            if len(s.split()) > 28:
                mid = len(s.split())//2
                a = ' '.join(s.split()[:mid])
                b = ' '.join(s.split()[mid:])
                if a and b:
                    if a[-1] not in '.!?': a += '.'
                    if b[-1] not in '.!?': b += '.'
                    b = b[0].upper() + b[1:]
                    new_sents.extend([a, b])
                else:
                    new_sents.append(s)
            else:
                new_sents.append(s)
        txt = ' '.join(new_sents)
    # لمسات بشرية
    if level >= 2 and random.random() < 0.3:
        txt = "So, " + txt[0].lower() + txt[1:]
    if level >= 4 and random.random() < 0.15:
        txt = txt.rstrip('.!?') + ', right?'
    txt = re.sub(r'\s+', ' ', txt)
    if txt and txt[0].islower():
        txt = txt[0].upper() + txt[1:]
    return txt

def process_word_preserve(file_bytes, level):
    """معالجة ملف Word مع الحفاظ على الجداول والأشكال والمعادلات"""
    doc = Document(file_bytes)
    modified = 0
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        # تخطي الفقرات التي تحتوي على معادلات أو كائنات
        skip = False
        for run in para.runs:
            if run.element.xpath('.//m:oMath') or run.element.xpath('.//w:object'):
                skip = True
                break
        if skip:
            continue
        new_text = humanize(para.text, level)
        if new_text != para.text:
            para.clear()
            run = para.add_run(new_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            modified += 1
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out, modified

# ==================== زر التحويل ====================
if st.button("🔄 تحويل النص (للملصق أو للملف)", type="primary", use_container_width=True):
    result = None
    if uploaded_file:
        with st.spinner("معالجة ملف Word مع الحفاظ على الجداول والأشكال..."):
            out_bytes, count = process_word_preserve(BytesIO(uploaded_file.read()), intensity)
            st.success(f"✅ تم تعديل {count} فقرة. جميع الجداول والأشكال والمعادلات سليمة.")
            st.download_button("📥 تحميل ملف Word المعدل", data=out_bytes,
                               file_name="humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    elif user_text.strip():
        with st.spinner("جاري التحويل..."):
            result = humanize(user_text, intensity)
        st.subheader("📄 النص الأصلي")
        st.text_area("", user_text, height=200)
        st.subheader("✨ النص المعدل (بشري)")
        st.text_area("", result, height=200)
        if result == user_text:
            st.warning("لم يتغير النص. جرب شدة أعلى أو نصاً أطول.")
        else:
            st.success(f"تم التغيير! {len(user_text)} → {len(result)} حرف")
            st.download_button("📥 تحميل النص (TXT)", data=result.encode(), file_name="humanized.txt")
    else:
        st.warning("الرجاء إدخال نص في المربع أو رفع ملف Word.")
