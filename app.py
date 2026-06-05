#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Precision Human Rewriter
يحافظ على المصطلحات التقنية والاستشهادات، فقط يستبدل الكلمات المحظورة ويعيد صياغة الجمل الطويلة قليلاً.
"""

import html
import re
import random
from io import BytesIO
from typing import List, Optional

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

try:
    import docx2txt
    import pypdf
    HAS_EXTRACT = True
except ImportError:
    HAS_EXTRACT = False

st.set_page_config(page_title="DeepClean Studio - Precision", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ============================================================
# قاموس الكلمات المحظورة فقط (لا نغير العبارات الكاملة)
# ============================================================
BANNED_TO_HUMAN = {
    # الأفعال والعبارات الشائعة (نستبدل فقط الكلمات وليس العبارات الكاملة)
    "constitute": "are", "constituting": "making up", "constitutes": "is",
    "forecasts": "expects", "forecast": "project", "forecasting": "prediction",
    "cumulative": "total", "exceeding": "above", "exceed": "go above",
    "trajectories": "paths", "committed": "plans", "generating": "producing",
    "utility-scale": "large", "dominant share": "most", "routinely": "often",
    "pronounced": "noticeable", "uniquely": "", "difficult": "hard",
    "fragmented": "disconnected", "expose": "provide", "derived": "calculated",
    "viability": "practicality", "fragmentation": "separation",
    "systematically": "often", "cross-domain": "interdisciplinary",
    "my last training update": "", "as an ai": "", "i hope this helps": "",
    "additionally": "also", "moreover": "plus", "furthermore": "then",
    "consequently": "so", "hence": "thus", "crucial": "very important",
    "pivotal": "key", "vital": "essential", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "broad",
    "delve": "explore", "showcase": "show", "underscore": "emphasize",
    "highlight": "point out", "resonate": "connect with", "align with": "match",
    "garner": "receive", "tapestry": "range", "testament": "evidence",
    "landscape": "situation", "intricate": "complex", "multifaceted": "many-sided",
    "serve as": "act as", "stands as": "is", "marks a turning point": "represents a change",
    "sets the stage for": "prepares for", "plays a key role": "helps",
    "in conclusion": "", "in summary": "", "overall": "generally",
    "it is important to note": "note that", "not only": "", "but also": "and also",
    "uniquely": "", "this article will": "", "we will": "I will",
    "key takeaways": "main points", "in today's world": "currently",
    "plays a crucial role": "is important in",
}

# ============================================================
# المحرك الدقيق - يحافظ على البنية والاستشهادات
# ============================================================
class PrecisionHumanRewriter:
    def __init__(self, text: str, seed: int = 42):
        self.text = text
        random.seed(seed)

    def _replace_banned_words_only(self, text: str) -> str:
        """استبدال الكلمات المحظورة فقط، دون تغيير بنية الجمل."""
        # ترتيب حسب الطول (الأطول أولاً) لتجنب الاستبدال الجزئي
        sorted_terms = sorted(BANNED_TO_HUMAN.keys(), key=len, reverse=True)
        for old in sorted_terms:
            new = BANNED_TO_HUMAN[old]
            if new == "":
                # حذف الكلمة نهائياً مع المسافات الزائدة
                text = re.sub(rf'\s*\b{re.escape(old)}\b\s*', ' ', text, flags=re.I)
            else:
                # استبدال مع الحفاظ على حالة الحرف الأول
                def replacer(match):
                    matched = match.group(0)
                    if matched[0].isupper():
                        return new[0].upper() + new[1:]
                    return new
                text = re.sub(rf'\b{re.escape(old)}\b', replacer, text, flags=re.I)
        # تنظيف المسافات الزائدة الناتجة عن الحذف
        text = re.sub(r'\s+', ' ', text)
        return text

    def _clean_punctuation(self, text: str) -> str:
        """تنظيف علامات الترقيم دون تغيير الاستشهادات."""
        # حماية الاستشهادات مؤقتاً
        citations = []
        def protect_cit(match):
            citations.append(match.group(0))
            return f"__CIT_{len(citations)-1}__"
        
        text = re.sub(r'\[\d+(?:[-,;]\s*\d+)*\]', protect_cit, text)
        text = re.sub(r'\([^)]*\d{4}[^)]*\)', protect_cit, text)
        
        # تنظيف الترقيم
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\.{2,}', '.', text)
        text = re.sub(r'\s+,', ',', text)
        text = re.sub(r'([.!?])\s+([a-z])', lambda m: f'{m.group(1)} {m.group(2).upper()}', text)
        
        # إعادة الاستشهادات
        for i, cit in enumerate(citations):
            text = text.replace(f"__CIT_{i}__", cit)
        
        return text

    def _fix_em_dashes(self, text: str) -> str:
        """تحويل الشرطات الطويلة إلى فواصل أو نقاط."""
        # استبدال الشرطات الطويلة بفواصل ونقاط
        text = text.replace('—', ', ')
        text = text.replace('–', '-')
        # تنظيف المسافات حول الشرطات
        text = re.sub(r'\s*[-]\s*', ' - ', text)
        return text

    def _preserve_citation_positions(self, original: str, rewritten: str) -> str:
        """ضمان بقاء الاستشهادات في مواقعها الأصلية."""
        # استخراج جميع الاستشهادات من النص الأصلي مع مواقعها
        cit_pattern = re.compile(r'(\[\d+(?:[-,;]\s*\d+)*\]|\([^)]*\d{4}[^)]*\))')
        orig_cits = cit_pattern.findall(original)
        
        if not orig_cits:
            return rewritten
        
        # إزالة جميع الاستشهادات من النص المعدل
        rewritten = cit_pattern.sub('', rewritten)
        
        # تقسيم النص المعدل إلى جمل
        sentences = re.split(r'(?<=[.!?])\s+', rewritten)
        
        # توزيع الاستشهادات على الجمل (واحد تقريباً لكل جملة)
        cit_index = 0
        new_sentences = []
        for sent in sentences:
            if cit_index < len(orig_cits):
                # إضافة استشهاد إلى هذه الجملة إذا كانت تحتوي على معنى مكتمل
                if len(sent.split()) > 8:
                    # نضيف الاستشهاد قبل النقطة النهائية
                    sent = sent.rstrip('.!?') + ' ' + orig_cits[cit_index] + '.'
                    cit_index += 1
            new_sentences.append(sent)
        
        return ' '.join(new_sentences)

    def _split_overly_long_sentences(self, text: str) -> str:
        """تقسيم الجمل الطويلة جداً (أكثر من 35 كلمة) إلى جملتين فقط."""
        sentences = re.split(r'(?<=[.!?])\s+', text)
        new_sentences = []
        
        for sent in sentences:
            word_count = len(sent.split())
            if word_count > 35:
                # نبحث عن فاصلة أو "and" أو "but" للتقسيم
                for sep in [', ', ' and ', ' but ', ' so ']:
                    if sep in sent:
                        parts = sent.split(sep, 1)
                        if len(parts) == 2 and parts[0] and parts[1]:
                            first = parts[0].strip()
                            second = parts[1].strip()
                            if first and second:
                                if first[-1] not in '.!?':
                                    first += '.'
                                second = second[0].upper() + second[1:]
                                if second[-1] not in '.!?':
                                    second += '.'
                                new_sentences.extend([first, second])
                                break
                else:
                    # تقسيم عادي في المنتصف
                    words = sent.split()
                    mid = len(words) // 2
                    first = ' '.join(words[:mid]).strip()
                    second = ' '.join(words[mid:]).strip()
                    if first and second:
                        if first[-1] not in '.!?':
                            first += '.'
                        second = second[0].upper() + second[1:]
                        if second[-1] not in '.!?':
                            second += '.'
                        new_sentences.extend([first, second])
            else:
                new_sentences.append(sent)
        
        return ' '.join(new_sentences)

    def _add_minimal_human_noise(self, text: str) -> str:
        """إضافة لمسة بشرية واحدة فقط (بداية واحدة) دون تشويه."""
        # نضيف "So" في البداية فقط إذا كان النص طويلاً
        if len(text) > 100 and not text.lower().startswith(('so', 'well', 'now')):
            return 'So, ' + text[0].lower() + text[1:]
        return text

    def run(self) -> str:
        text = self.text
        
        # إزالة بقايا الروبوتات
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is|as an ai|as a large language model).*\n', '', text, flags=re.M)
        text = re.sub(r'\*+', '', text)
        text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
        
        # استبدال الكلمات المحظورة فقط
        text = self._replace_banned_words_only(text)
        
        # تحويل الشرطات الطويلة
        text = self._fix_em_dashes(text)
        
        # تقسيم الجمل الطويلة جداً فقط
        text = self._split_overly_long_sentences(text)
        
        # إضافة لمسة بشرية بسيطة
        text = self._add_minimal_human_noise(text)
        
        # تنظيف الترقيم
        text = self._clean_punctuation(text)
        
        # إعادة بناء الاستشهادات في مواقعها
        text = self._preserve_citation_positions(self.text, text)
        
        # تنظيف نهائي
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        return text


# ============================================================
# دوال مساعدة
# ============================================================
def simple_token_count(text: str) -> int:
    return len(re.findall(r'\b\w+\b', text))

def extract_file_text(uploaded_file) -> str:
    if not HAS_EXTRACT:
        return "الرجاء تثبيت المكتبات: pip install docx2txt pypdf"
    name = uploaded_file.name.lower()
    if name.endswith('.txt'):
        try:
            return uploaded_file.read().decode('utf-8')
        except:
            return uploaded_file.read().decode('utf-8-sig')
    elif name.endswith('.docx'):
        return docx2txt.process(uploaded_file) or ''
    elif name.endswith('.pdf'):
        reader = pypdf.PdfReader(uploaded_file)
        return '\n'.join(page.extract_text() or '' for page in reader.pages)
    return ''

def make_word_file(text: str) -> BytesIO:
    doc = Document()
    doc.core_properties.author = AUTHOR_NAME
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
    
    for line in text.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
        else:
            doc.add_paragraph()
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def html_preview(text: str) -> str:
    return f"<div style='font-family: Times New Roman; font-size: 12pt; line-height: 1.4;'>{html.escape(text).replace(chr(10), '<br>')}</div>"


# ============================================================
# واجهة Streamlit
# ============================================================
def main():
    st.title("✍️ DeepClean Studio – Precision Human Rewriter")
    st.caption("يحافظ على المصطلحات التقنية والاستشهادات، يستبدل الكلمات المحظورة فقط")
    st.caption(AUTHOR_NAME)
    
    with st.sidebar:
        st.header("الإعدادات")
        source = st.radio("المصدر", ["📄 لصق نص", "📁 رفع ملف"])
        user_text = ""
        if source == "📁 رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                user_text = extract_file_text(uploaded)
                if user_text:
                    st.success("تم تحميل الملف")
        else:
            user_text = st.text_area("ألصق النص الأكاديمي هنا", height=250)
        
        seed = st.number_input("بذرة عشوائية", value=42, step=1)
        
        if st.button("🚀 إعادة الكتابة", type="primary", use_container_width=True):
            if not user_text:
                st.warning("الرجاء إدخال نص أو رفع ملف.")
            else:
                with st.spinner("جاري إعادة الكتابة..."):
                    engine = PrecisionHumanRewriter(user_text, seed=seed)
                    result = engine.run()
                    st.session_state.result = result
                    st.session_state.original = user_text
                    st.session_state.done = True
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📄 النص الأصلي")
        if user_text:
            st.text_area("", user_text, height=450, key="orig_area")
            st.caption(f"عدد الكلمات: {simple_token_count(user_text)}")
        else:
            st.info("أدخل نصًا من الشريط الجانبي.")
    
    with col2:
        st.subheader("✨ النص المعدل")
        if st.session_state.get("done") and st.session_state.get("result"):
            res = st.session_state.result
            st.markdown(html_preview(res), unsafe_allow_html=True)
            st.text_area("", res, height=450, key="res_area", label_visibility="collapsed")
            st.caption(f"عدد الكلمات: {simple_token_count(res)}")
            docx_file = make_word_file(res)
            st.download_button("📥 تحميل Word", data=docx_file, file_name="precision_humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        else:
            st.info("ستظهر النسخة المعدلة هنا بعد المعالجة.")

if __name__ == "__main__":
    if "done" not in st.session_state:
        st.session_state.done = False
    if "result" not in st.session_state:
        st.session_state.result = ""
    if "original" not in st.session_state:
        st.session_state.original = ""
    main()
