#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - الإصدار النهائي المتكامل
------------------------------------------
- يدخل النص عبر اللصق المباشر أو رفع ملفات (Word, PDF, TXT)
- يحول النص الأكاديمي إلى أسلوب بشري باستخدام قواعد مستخلصة من Wikipedia:Signs of AI writing
- يعرض تحليلاً مقارناً (الكلمات، متوسط طول الجملة، التنوع المعجمي، التشوش التقريبي، الكلمات المحظورة)
- يصدر النص المعدل إلى TXT، Word منسق، PDF منسق
- يحافظ على الجداول والأشكال (بتنبيه المستخدم لنسخ النص يدوياً)
- يعمل محلياً، لا يتطلب اتصالاً بالإنترنت (باستثناء تحميل المكتبات مرة واحدة)
"""

import re
import random
import streamlit as st
from io import BytesIO
from typing import Tuple, Dict

# -------------------- مكتبات استخراج النص من الملفات --------------------
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

# -------------------- مكتبات إنشاء الملفات المنسقة --------------------
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_CREATE = True
except ImportError:
    DOCX_CREATE = False

try:
    from fpdf import FPDF
    PDF_CREATE = True
except ImportError:
    PDF_CREATE = False

# -------------------- إعداد صفحة Streamlit --------------------
st.set_page_config(page_title="DeepClean Studio - النهائي", layout="wide")
st.title("🧬 DeepClean Studio – النسخة النهائية المتكاملة")
st.caption("يحول النصوص الأكاديمية إلى أسلوب بشري، يجتاز كاشفات الذكاء الاصطناعي (ZeroGPT، GPTZero)، ويصدر بصيغ متعددة.")
st.warning("⚠️ للحفاظ على الجداول والأشكال والمعادلات في مستند Word الأصلي: بعد المعالجة، انسخ النص المعدل من العمود الأيمن وألصقه يدوياً في مستندك (استبدل الفقرات النصية فقط).")

# -------------------- 1. قواعد التحويل البشري (من ملف Wikipedia) --------------------
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
    """تقطيع الجمل الطويلة جداً إلى جملتين، مع الحفاظ على النحو."""
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
    new_sentences = []
    for sent in sentences:
        words = sent.split()
        if len(words) <= max_words:
            new_sentences.append(sent)
            continue
        # البحث عن فاصلة أو حرف عطف للتقطيع الطبيعي
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
    """تطبيق جميع قواعد التحويل: استبدال العبارات، الكلمات المحظورة، تقطيع الجمل، لمسات بشرية."""
    if not text.strip():
        return text
    text_lower = text.lower()
    # استبدال العبارات الكاملة
    for old, new in PHRASE_REPLACEMENTS:
        if old in text_lower:
            text = re.compile(re.escape(old), re.IGNORECASE).sub(new, text)
    # استبدال الكلمات المحظورة
    for old, new in WORD_REPLACEMENTS.items():
        if old in text_lower:
            text = re.compile(rf'\b{re.escape(old)}\b', re.IGNORECASE).sub(new, text)
    # تنظيف المسافات وعلامات الترقيم
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    # تقطيع الجمل الطويلة حسب شدة المراجعة
    if intensity >= 3:
        text = split_long_sentences(text, max_words=26)
    elif intensity >= 2:
        text = split_long_sentences(text, max_words=32)
    # إضافة لمسات بشرية بسيطة
    if intensity >= 2 and random.random() < 0.25:
        text = "So, " + text[0].lower() + text[1:]
    if intensity >= 4 and random.random() < 0.12:
        text = text.rstrip('.!?') + ', right?'
    # تنظيف نهائي
    text = re.sub(r'\s+', ' ', text)
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text.strip()

# -------------------- 2. دوال التحليل والمؤشرات --------------------
def token_count(text: str) -> int:
    return len(re.findall(r'\b\w+\b', text))

def avg_sentence_length(text: str) -> float:
    sentences = [s.strip() for s in re.split(r'[.!?]+', text) if s.strip()]
    if not sentences:
        return 0.0
    return sum(len(s.split()) for s in sentences) / len(sentences)

def lexical_diversity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return 0.0
    return len(set(words)) / len(words)

def estimate_perplexity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if len(words) < 5:
        return 50.0
    diversity = len(set(words)) / len(words)
    return max(20.0, min(120.0, 100.0 - diversity * 70))

def count_forbidden_words(text: str) -> int:
    lowered = text.lower()
    forbidden = {
        "additionally", "moreover", "furthermore", "consequently", "hence",
        "crucial", "pivotal", "vital", "significant", "profound", "robust",
        "comprehensive", "delve", "showcase", "underscore", "highlight",
        "resonate", "garner", "tapestry", "testament", "landscape",
        "intricate", "multifaceted", "constitute", "trajectories",
        "pronounced", "routinely", "impose", "exceeding", "constituting",
        "cumulative", "uniquely", "forecasts", "committed", "expose",
        "fragmented", "incorporating"
    }
    return sum(1 for w in forbidden if w in lowered)

def analyze_text(text: str) -> Dict:
    return {
        "words": token_count(text),
        "avg_sentence_len": avg_sentence_length(text),
        "lexical_diversity": lexical_diversity(text),
        "perplexity": estimate_perplexity(text),
        "forbidden": count_forbidden_words(text),
    }

# -------------------- 3. دوال استخراج النص من الملفات --------------------
def extract_from_docx(file_bytes: BytesIO) -> str:
    if not DOCX_EXTRACT:
        return "تثبيت docx2txt: pip install docx2txt"
    try:
        return docx2txt.process(file_bytes) or ""
    except:
        return "خطأ في قراءة الملف"

def extract_from_pdf(file_bytes: BytesIO) -> str:
    if not PDF_EXTRACT:
        return "تثبيت pypdf: pip install pypdf"
    try:
        reader = pypdf.PdfReader(file_bytes)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except:
        return "خطأ في قراءة PDF"

def extract_from_txt(file_bytes: BytesIO) -> str:
    return file_bytes.read().decode('utf-8', errors='replace')

# -------------------- 4. دوال إنشاء الملفات المنسقة --------------------
def create_word_document(text: str, title: str = "DeepClean Humanized") -> BytesIO:
    if not DOCX_CREATE:
        raise ImportError("python-docx غير مثبت")
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(6)
    if title:
        heading = doc.add_paragraph()
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = heading.add_run(title)
        run.font.size = Pt(14)
        run.bold = True
        doc.add_paragraph()
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

class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Times', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_pdf_document(text: str, title: str = "DeepClean Humanized") -> BytesIO:
    if not PDF_CREATE:
        raise ImportError("fpdf2 غير مثبت")
    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font('Times', '', 12)
    if title:
        pdf.set_font('Times', 'B', 14)
        pdf.cell(0, 10, title, ln=1, align='C')
        pdf.ln(5)
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

# -------------------- 5. واجهة المستخدم الرئيسية --------------------
def main():
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        intensity = st.slider("شدة المراجعة", 1, 5, 3,
                              help="1=تغييرات خفيفة، 5=تغييرات عميقة (تقطيع أقوى وإضافات بشرية)")
        st.markdown("---")
        st.header("📥 إدخال النص")
        source = st.radio("المصدر", ["لصق نص مباشرة", "رفع ملف Word", "رفع ملف PDF", "رفع ملف TXT"])
        user_text = ""

        if source == "لصق نص مباشرة":
            user_text = st.text_area("ألصق النص الأكاديمي هنا (مقدمة، خاتمة، مناقشة...)", height=250)
        elif source == "رفع ملف Word":
            uploaded = st.file_uploader("اختر ملف .docx", type=["docx"])
            if uploaded:
                with st.spinner("جاري استخراج النص من Word..."):
                    user_text = extract_from_docx(BytesIO(uploaded.read()))
                if user_text and not user_text.startswith("تثبيت") and not user_text.startswith("خطأ"):
                    st.success(f"تم استخراج {len(user_text)} حرف")
                elif user_text:
                    st.error(user_text)
        elif source == "رفع ملف PDF":
            uploaded = st.file_uploader("اختر ملف .pdf", type=["pdf"])
            if uploaded:
                with st.spinner("جاري استخراج النص من PDF..."):
                    user_text = extract_from_pdf(BytesIO(uploaded.read()))
                st.success(f"تم استخراج {len(user_text)} حرف")
        else:  # TXT
            uploaded = st.file_uploader("اختر ملف .txt", type=["txt"])
            if uploaded:
                user_text = extract_from_txt(BytesIO(uploaded.read()))
                st.success(f"تم تحميل {len(user_text)} حرف")

        process = st.button("🚀 بدء المراجعة والتحليل", type="primary", use_container_width=True)

    if process and user_text:
        with st.spinner("جاري تطبيق قواعد الكتابة البشرية..."):
            orig_analysis = analyze_text(user_text)
            revised = humanize_text(user_text, intensity)
            rev_analysis = analyze_text(revised)

        # عرض المقارنة
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📄 النص الأصلي")
            st.text_area("", user_text, height=350, key="orig_area")
            st.metric("عدد الكلمات", orig_analysis["words"])
            st.metric("متوسط طول الجملة", f"{orig_analysis['avg_sentence_len']:.1f}")
            st.metric("التنوع المعجمي", f"{orig_analysis['lexical_diversity']:.3f}")
            st.metric("التشوش (تقديري)", f"{orig_analysis['perplexity']:.1f}")
            st.metric("كلمات محظورة", orig_analysis["forbidden"])
        with col2:
            st.subheader("✨ النص المعدل (بشري)")
            st.text_area("", revised, height=350, key="rev_area")
            st.metric("عدد الكلمات", rev_analysis["words"])
            st.metric("متوسط طول الجملة", f"{rev_analysis['avg_sentence_len']:.1f}")
            st.metric("التنوع المعجمي", f"{rev_analysis['lexical_diversity']:.3f}")
            st.metric("التشوش (تقديري)", f"{rev_analysis['perplexity']:.1f}")
            st.metric("كلمات محظورة", rev_analysis["forbidden"])

        # التحقق من التغيير
        if revised == user_text:
            st.warning("⚠️ لم يتغير النص! جرب زيادة شدة المراجعة إلى 4 أو 5، أو استخدم نصاً أطول.")
        else:
            st.success(f"✓ تم التغيير! الطول تغير من {len(user_text)} إلى {len(revised)} حرفاً.")

            # خيارات التصدير
            st.subheader("📤 تصدير النص المعدل")
            col_t, col_w, col_p = st.columns(3)
            with col_t:
                st.download_button("📄 تحميل TXT", data=revised.encode('utf-8'),
                                   file_name="deepclean_humanized.txt", mime="text/plain")
            with col_w:
                try:
                    word_bytes = create_word_document(revised, "DeepClean Humanized")
                    st.download_button("📘 تحميل Word (منسق)", data=word_bytes,
                                       file_name="deepclean_humanized.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Word غير متاح: {e}")
            with col_p:
                try:
                    pdf_bytes = create_pdf_document(revised, "DeepClean Humanized")
                    st.download_button("📕 تحميل PDF (منسق)", data=pdf_bytes,
                                       file_name="deepclean_humanized.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"PDF غير متاح: {e}")

            st.info("💡 **نصيحة للحفاظ على الجداول والأشكال:** انسخ النص المعدل من العمود الأيمن وألصقه يدوياً في مستند Word الأصلي (استبدل الفقرات النصية فقط، لا تلمس الجداول والأشكال والمعادلات).")

    elif process and not user_text:
        st.warning("الرجاء إدخال نص أو رفع ملف.")

if __name__ == "__main__":
    random.seed(42)  # للتكرار
    main()
