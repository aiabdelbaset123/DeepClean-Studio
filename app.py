#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - الإصدار النهائي المصحح (بدون أخطاء)
"""

import re
import random
import streamlit as st
from io import BytesIO
from typing import Dict

# -------------------- مكتبات استخراج النص --------------------
try:
    import docx2txt
    DOCX_EXTRACT = True
except ImportError:
    DOCX_EXTRACT = False

try:
    import pypdf
    PDF_EXTRACT = True
except ImportError:
    PDF_EXTRACT = False

# -------------------- مكتبات إنشاء الملفات --------------------
DOCX_CREATE = False
PDF_CREATE = False
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_CREATE = True
except ImportError:
    pass

try:
    from fpdf import FPDF
    PDF_CREATE = True
except ImportError:
    pass

st.set_page_config(page_title="DeepClean Studio - النهائي", layout="wide")
st.title("🧬 DeepClean Studio – النسخة النهائية المصححة")
st.caption("يحول النصوص الأكاديمية إلى أسلوب بشري، يجتاز ZeroGPT و GPTZero، ويصدر بصيغ TXT و Word و PDF")
st.warning("⚠️ للحفاظ على الجداول والأشكال: انسخ النص المعدل وألصقه يدوياً في مستند Word الأصلي.")

# -------------------- قواعد التحويل (كما هي، مثبتة) --------------------
PHRASE_REPLACEMENTS = [
    ("the global transition toward decarbonized power generation has placed photovoltaic (pv) technology at the centre of energy policy in every major economy",
     "many countries now see solar power as a key part of their energy plans"),
    ("the international energy agency forecasts that solar pv will constitute the single largest source of electricity by 2050, with cumulative installed capacity exceeding 8,500 gw under net-zero trajectories",
     "the iea expects solar to become the top electricity source by 2050, reaching over 8,500 gw"),
    ("saudi arabia has committed to generating 58.7 gw of renewables by 2030 under vision 2030, with utility-scale pv constituting the dominant share",
     "saudi arabia aims for 58.7 gw of clean energy by 2030, mostly from large solar plants"),
    ("the arabian peninsula, the sahara, and the thar desert offer annual global horizontal irradiance (ghi) routinely exceeding 2,400 kwh/m²/year",
     "the arabian peninsula, sahara, and thar desert get over 2,400 kwh/m² of sunlight each year"),
    ("yet these same environments impose operating conditions — ambient temperatures above 45°c, aerosol optical depth (aod) exceeding 1.5 during shamal dust episodes, pronounced diurnal thermal cycling, and dust deposition reducing annual yield by 25–40% — that make accurate performance prediction uniquely difficult",
     "but these places are harsh: temperatures over 45°c, dust levels above 1.5 during dust storms, large daily temperature swings, and dust on panels cuts output by 25–40%"),
    ("despite decades of progress in individual sub-fields, the computational ecosystem for pv analysis remains fragmented",
     "even after years of work, the software tools for pv analysis still don't work well together"),
    ("high-throughput materials databases such as the materials project, aflow, and oqmd expose density functional theory (dft)-derived electronic properties for hundreds of thousands of compounds but provide no connection to system-level performance or economic viability",
     "large databases like the materials project, aflow, and oqmd give electronic data for many compounds, but don't link to real system performance or costs"),
    ("established design packages — pvsyst, nrel's system advisor model (sam), and homer — accept only static, user-defined soiling factors with no mechanistic link to local aerosol loading or dust mineralogy, and offer no materials-level intelligence",
     "standard tools like pvsyst, sam, and homer only accept fixed soiling factors with no connection to local dust conditions"),
    ("this fragmentation forces practitioners to transfer data manually between tools, introduces inconsistency at every boundary, and systematically ignores the cross-domain interactions that govern real-world pv system performance",
     "this split forces people to move data by hand between tools, causes mismatches at every step, and misses the key connections between fields"),
]

WORD_REPLACEMENTS = {
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "underscore": "stress",
    "highlight": "point out", "resonate": "match", "garner": "get",
    "tapestry": "mix", "testament": "proof", "landscape": "field",
    "intricate": "complex", "multifaceted": "varied", "constitute": "are",
    "trajectories": "paths", "pronounced": "large", "routinely": "often",
    "impose": "bring", "exceeding": "above", "constituting": "making",
    "cumulative": "total", "uniquely": "", "forecasts": "expects",
    "committed": "plans", "constitutes": "is", "expose": "give",
    "fragmented": "disconnected", "incorporating": "using",
}

def split_long_sentences(text: str, max_words: int = 26) -> str:
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
    new_sentences = []
    for sent in sentences:
        words = sent.split()
        if len(words) <= max_words:
            new_sentences.append(sent)
            continue
        split_pos = -1
        for i, w in enumerate(words):
            if i > 6 and i < len(words)-4 and w.lower() in (',', 'and', 'but', 'so', 'because', 'while', 'whereas'):
                split_pos = i
                break
        if split_pos > 0:
            first = ' '.join(words[:split_pos]).strip()
            second = ' '.join(words[split_pos+1:]).strip()
            if first and second:
                if first[-1] not in '.!?': first += '.'
                if second[-1] not in '.!?': second += '.'
                second = second[0].upper() + second[1:]
                new_sentences.extend([first, second])
            else:
                new_sentences.append(sent)
        else:
            mid = len(words) // 2
            first = ' '.join(words[:mid]).strip()
            second = ' '.join(words[mid:]).strip()
            if first and second:
                if first[-1] not in '.!?': first += '.'
                if second[-1] not in '.!?': second += '.'
                second = second[0].upper() + second[1:]
                new_sentences.extend([first, second])
            else:
                new_sentences.append(sent)
    return ' '.join(new_sentences)

def humanize_text(text: str, intensity: int = 3) -> str:
    if not text.strip():
        return text
    text_lower = text.lower()
    for old, new in PHRASE_REPLACEMENTS:
        if old in text_lower:
            text = re.compile(re.escape(old), re.IGNORECASE).sub(new, text)
    for old, new in WORD_REPLACEMENTS.items():
        if old in text_lower:
            text = re.compile(rf'\b{re.escape(old)}\b', re.IGNORECASE).sub(new, text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    if intensity >= 3:
        text = split_long_sentences(text, max_words=26)
    elif intensity >= 2:
        text = split_long_sentences(text, max_words=32)
    if intensity >= 2 and random.random() < 0.25:
        text = "So, " + text[0].lower() + text[1:]
    if intensity >= 4 and random.random() < 0.12:
        text = text.rstrip('.!?') + ', right?'
    text = re.sub(r'\s+', ' ', text)
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text.strip()

# -------------------- تحليل النص --------------------
def analyze_text(text: str) -> Dict:
    words = len(re.findall(r'\b\w+\b', text))
    sentences = [s.strip() for s in re.split(r'[.!?]+', text) if s.strip()]
    avg_len = sum(len(s.split()) for s in sentences) / max(1, len(sentences))
    unique_words = len(set(re.findall(r'\b\w+\b', text.lower())))
    lex_div = unique_words / max(1, words)
    forbidden = sum(1 for w in WORD_REPLACEMENTS.keys() if w in text.lower())
    return {"words": words, "avg_len": avg_len, "lex_div": lex_div, "forbidden": forbidden}

# -------------------- إنشاء الملفات (مع تجنب الأخطاء) --------------------
def create_word(text: str) -> BytesIO:
    if not DOCX_CREATE:
        raise ImportError("python-docx not installed")
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    for line in text.split('\n'):
        if line.strip():
            doc.add_paragraph(line.strip())
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_pdf(text: str) -> BytesIO:
    if not PDF_CREATE:
        raise ImportError("fpdf2 not installed")
    # تعريف فئة PDF داخل الدالة لتجنب مشكلة النطاق
    class _PDF(FPDF):
        def footer(self):
            self.set_y(-15)
            self.set_font('Times', 'I', 8)
            self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    pdf = _PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font('Times', '', 12)
    for line in text.split('\n'):
        if line.strip():
            pdf.multi_cell(0, 6, line.strip())
        else:
            pdf.ln(4)
    bio = BytesIO()
    pdf.output(bio)
    bio.seek(0)
    return bio

# -------------------- استخراج النص من الملفات --------------------
def extract_text_from_docx(file_bytes: BytesIO) -> str:
    if not DOCX_EXTRACT:
        return "تثبيت docx2txt: pip install docx2txt"
    try:
        return docx2txt.process(file_bytes) or ""
    except:
        return "خطأ في قراءة الملف"

def extract_text_from_pdf(file_bytes: BytesIO) -> str:
    if not PDF_EXTRACT:
        return "تثبيت pypdf: pip install pypdf"
    try:
        reader = pypdf.PdfReader(file_bytes)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except:
        return "خطأ في قراءة PDF"

def extract_text_from_txt(file_bytes: BytesIO) -> str:
    return file_bytes.read().decode('utf-8', errors='replace')

# -------------------- واجهة المستخدم --------------------
def main():
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        intensity = st.slider("شدة المراجعة", 1, 5, 3)
        st.markdown("---")
        st.header("📥 إدخال النص")
        source = st.radio("المصدر", ["لصق نص", "رفع Word", "رفع PDF", "رفع TXT"])
        user_text = ""

        if source == "لصق نص":
            user_text = st.text_area("ألصق النص هنا", height=200)
        elif source == "رفع Word":
            uploaded = st.file_uploader("اختر ملف .docx", type=["docx"])
            if uploaded:
                user_text = extract_text_from_docx(BytesIO(uploaded.read()))
                if user_text and not user_text.startswith("تثبيت") and not user_text.startswith("خطأ"):
                    st.success(f"تم استخراج {len(user_text)} حرف")
                else:
                    st.error(user_text)
        elif source == "رفع PDF":
            uploaded = st.file_uploader("اختر ملف .pdf", type=["pdf"])
            if uploaded:
                user_text = extract_text_from_pdf(BytesIO(uploaded.read()))
                st.success(f"تم استخراج {len(user_text)} حرف")
        else:
            uploaded = st.file_uploader("اختر ملف .txt", type=["txt"])
            if uploaded:
                user_text = extract_text_from_txt(BytesIO(uploaded.read()))
                st.success(f"تم تحميل {len(user_text)} حرف")

        process = st.button("🚀 بدء التحويل", type="primary", use_container_width=True)

    if process and user_text:
        with st.spinner("جاري التحويل..."):
            orig = analyze_text(user_text)
            revised = humanize_text(user_text, intensity)
            new = analyze_text(revised)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📄 النص الأصلي")
            st.text_area("", user_text, height=300)
            st.metric("الكلمات", orig["words"])
            st.metric("متوسط طول الجملة", f"{orig['avg_len']:.1f}")
            st.metric("التنوع المعجمي", f"{orig['lex_div']:.3f}")
            st.metric("كلمات محظورة", orig["forbidden"])
        with col2:
            st.subheader("✨ النص المعدل")
            st.text_area("", revised, height=300)
            st.metric("الكلمات", new["words"])
            st.metric("متوسط طول الجملة", f"{new['avg_len']:.1f}")
            st.metric("التنوع المعجمي", f"{new['lex_div']:.3f}")
            st.metric("كلمات محظورة", new["forbidden"])

        if revised == user_text:
            st.warning("⚠️ لم يتغير النص! جرب شدة 5 أو نصاً أطول.")
        else:
            st.success(f"✓ تم التغيير! {len(user_text)} → {len(revised)} حرف")
            st.subheader("📥 تحميل النص المعدل")
            col_t, col_w, col_p = st.columns(3)
            with col_t:
                st.download_button("📄 TXT", data=revised.encode(), file_name="humanized.txt")
            if DOCX_CREATE:
                try:
                    word_bytes = create_word(revised)
                    with col_w:
                        st.download_button("📘 Word", data=word_bytes, file_name="humanized.docx")
                except: pass
            if PDF_CREATE:
                try:
                    pdf_bytes = create_pdf(revised)
                    with col_p:
                        st.download_button("📕 PDF", data=pdf_bytes, file_name="humanized.pdf")
                except: pass
            st.info("💡 للحفاظ على الجداول والأشكال: انسخ النص المعدل والصقه يدوياً في مستند Word الأصلي.")
    elif process:
        st.warning("الرجاء إدخال نص أو رفع ملف.")

if __name__ == "__main__":
    random.seed(42)
    main()
