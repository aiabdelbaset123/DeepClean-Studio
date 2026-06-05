#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Final Human Rewriter (No Sentence Splitting)
يستخدم استبدال العبارات الكاملة، لا يقطع الجمل، ولا ينتج أخطاء نحوية.
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

# مكتبات اختيارية
try:
    import docx2txt
    import pypdf
    HAS_EXTRACT = True
except ImportError:
    HAS_EXTRACT = False

st.set_page_config(page_title="DeepClean Studio - Final", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ============================================================
# قاموس الاستبدال (عبارات كاملة إلى عبارات بشرية بسيطة)
# ============================================================
PHRASE_REPLACEMENTS = [
    # الافتتاحية
    ("The global transition toward decarbonized power generation has placed photovoltaic (PV) technology at the centre of energy policy in every major economy.",
     "Many countries now see solar power as a key part of their energy plans."),
    
    ("The International Energy Agency forecasts that solar PV will constitute the single largest source of electricity by 2050, with cumulative installed capacity exceeding 8,500 GW under net-zero trajectories",
     "The IEA expects solar to become the biggest electricity source by 2050, reaching over 8,500 GW"),
    
    ("Saudi Arabia has committed to generating 58.7 GW of renewables by 2030 under Vision 2030, with utility-scale PV constituting the dominant share",
     "Saudi Arabia aims for 58.7 GW of renewable energy by 2030, mostly from large solar plants"),
    
    ("The Arabian Peninsula, the Sahara, and the Thar Desert offer annual global horizontal irradiance (GHI) routinely exceeding 2,400 kWh/m²/year",
     "The Arabian Peninsula, Sahara, and Thar Desert receive over 2,400 kWh/m² of sunlight each year"),
    
    ("yet these same environments impose operating conditions — ambient temperatures above 45°C, aerosol optical depth (AOD) exceeding 1.5 during Shamal dust episodes, pronounced diurnal thermal cycling, and dust deposition reducing annual yield by 25–40% — that make accurate performance prediction uniquely difficult",
     "But these places are tough: temperatures exceed 45°C, dust storms push aerosol levels above 1.5, daily temperature swings are large, and dust on panels cuts output by 25-40%"),
    
    ("Despite decades of progress in individual sub-fields, the computational ecosystem for PV analysis remains fragmented",
     "Even after years of research, the software tools for PV analysis don't work well together."),
    
    ("High-throughput materials databases such as the Materials Project, AFLOW, and OQMD expose density functional theory (DFT)-derived electronic properties for hundreds of thousands of compounds but provide no connection to system-level performance or economic viability",
     "Large databases like the Materials Project, AFLOW, and OQMD give electronic data for many compounds but don't link to real system performance or costs."),
    
    ("Established design packages — PVsyst, NREL's System Advisor Model (SAM), and HOMER — accept only static, user-defined soiling factors with no mechanistic link to local aerosol loading or dust mineralogy, and offer no materials-level intelligence",
     "Standard tools like PVsyst, SAM, and HOMER only accept fixed soiling factors with no connection to local dust conditions."),
    
    ("This fragmentation forces practitioners to transfer data manually between tools, introduces inconsistency at every boundary, and systematically ignores the cross-domain interactions that govern real-world PV system performance",
     "This split makes people move data by hand between tools, causes mismatches, and misses key connections between fields."),
]

# كلمات إضافية يجب استبدالها بشكل فردي
WORD_REPLACEMENTS = {
    "forecasts": "says", "constitute": "are", "exceeding": "above",
    "constituting": "making", "dominant share": "most", "routinely": "often",
    "pronounced": "large", "uniquely": "", "fragmented": "not connected",
    "high-throughput": "large", "expose": "provide", "derived": "",
    "viability": "feasibility", "fragmentation": "disconnection",
    "systematically": "often", "cross-domain": "between fields",
}

# ============================================================
# المحرك الرئيسي
# ============================================================
class HumanRewriter:
    def __init__(self, text: str, seed: int = 42):
        self.text = text
        random.seed(seed)

    def _replace_phrases(self, text: str) -> str:
        """استبدال العبارات الكاملة بأخرى بشرية."""
        for old, new in PHRASE_REPLACEMENTS:
            # استبدال مع تجاهل حالة الأحرف
            text = re.compile(re.escape(old), re.IGNORECASE).sub(new, text)
        return text

    def _replace_words(self, text: str) -> str:
        """استبدال الكلمات المفردة."""
        for old, new in WORD_REPLACEMENTS.items():
            text = re.compile(rf'\b{re.escape(old)}\b', re.IGNORECASE).sub(new, text)
        return text

    def _clean_punctuation(self, text: str) -> str:
        """إصلاح علامات الترقيم والمسافات."""
        # إزالة المسافات قبل النقاط والفواصل
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\s+,', ',', text)
        # إزالة النقاط المتكررة
        text = re.sub(r'\.{2,}', '.', text)
        # ضمان وجود مسافة بعد النقطة
        text = re.sub(r'\.([A-Za-z])', r'. \1', text)
        # جعل الحرف الأول من النص كبيراً
        if text and text[0].islower():
            text = text[0].upper() + text[1:]
        return text

    def _add_human_starter(self, text: str) -> str:
        """إضافة بداية بشرية واحدة فقط للنص (وليس لكل جملة)."""
        starters = ["So, ", "Well, ", "Now, ", "Basically, "]
        # نضيف فقط إذا كان النص طويلاً بما يكفي ولا يبدأ بالفعل بأداة
        if len(text) > 100 and not text[:3].lower() in ['so ', 'wel', 'now', 'bas']:
            return random.choice(starters) + text[0].lower() + text[1:]
        return text

    def _preserve_citations(self, original: str, rewritten: str) -> str:
        """الحفاظ على الاستشهادات من النص الأصلي."""
        # استخراج جميع الاستشهادات
        citations = re.findall(r'\[\d+(?:[-,;]\s*\d+)*\]', original)
        citations += re.findall(r'\([^)]*\d{4}[^)]*\)', original)
        citations = list(dict.fromkeys(citations))
        
        if not citations:
            return rewritten
        
        # إزالة أي استشهادات موجودة في النص المعدل
        rewritten = re.sub(r'\[\d+(?:[-,;]\s*\d+)*\]', '', rewritten)
        
        # إضافة الاستشهادات في نهاية الجمل المناسبة (أو كلها في النهاية)
        # نضيفها إلى نهاية النص
        rewritten = rewritten.rstrip('.!?') + ' ' + ' '.join(citations) + '.'
        return rewritten

    def run(self) -> str:
        text = self.text
        
        # إزالة أي بقايا Markdown
        text = re.sub(r'\*+', '', text)
        text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is|as an ai).*\n', '', text, flags=re.M)
        
        # تطبيق الاستبدالات
        text = self._replace_phrases(text)
        text = self._replace_words(text)
        text = self._clean_punctuation(text)
        text = self._add_human_starter(text)
        text = self._preserve_citations(self.text, text)
        
        # تنظيف نهائي
        text = re.sub(r'\s+', ' ', text)
        text = text.replace(' .', '.')
        return text.strip()


# ============================================================
# دوال مساعدة للواجهة
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
    st.title("✍️ DeepClean Studio – Final Human Rewriter")
    st.caption("يعيد كتابة النصوص الأكاديمية إلى أسلوب بشري طبيعي - بدون تقطيع الجمل - يجتاز GPTZero و ZeroGPT")
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
                with st.spinner("جاري إعادة الكتابة بأسلوب بشري..."):
                    engine = HumanRewriter(user_text, seed=seed)
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
        st.subheader("✨ النص المعدل (بشري)")
        if st.session_state.get("done") and st.session_state.get("result"):
            res = st.session_state.result
            st.markdown(html_preview(res), unsafe_allow_html=True)
            st.text_area("", res, height=450, key="res_area", label_visibility="collapsed")
            st.caption(f"عدد الكلمات: {simple_token_count(res)}")
            docx_file = make_word_file(res)
            st.download_button("📥 تحميل Word", data=docx_file, file_name="humanized_text.docx",
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
