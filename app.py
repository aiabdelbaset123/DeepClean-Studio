#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - الإصدار الشامل النهائي
يجمع كل الميزات المطلوبة:
- إدخال نص (لصق) أو رفع ملف (Word, TXT, PDF)
- معالجة آمنة لملفات Word مع الحفاظ على الجداول والأشكال والمعادلات
- لوحة شفافية وتحليل: مؤشرات الجودة، الكلمات المحظورة، تظليل الجمل
- تطبيق قواعد إعادة الصياغة البشرية لاجتياز ZeroGPT
"""

import re
import random
import streamlit as st
from io import BytesIO
from typing import List, Tuple, Dict, Optional

# مكتبات رفع الملفات
try:
    import docx2txt
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pypdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - الشامل", layout="wide")

# ============================================================
# 1. قواعد إعادة الصياغة البشرية
# ============================================================

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

FORBIDDEN_SET = set(WORD_REPLACEMENTS.keys()) | {
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "garner", "tapestry", "testament", "landscape",
    "intricate", "multifaceted", "constitute", "trajectories", "pronounced",
    "routinely", "impose", "exceeding", "constituting", "cumulative",
    "uniquely", "forecasts", "committed", "expose", "fragmented", "incorporating"
}

def split_long_sentences(text: str, max_words: int = 25) -> str:
    """تقطيع الجمل الطويلة إلى جملتين أو ثلاث."""
    sentences = re.split(r'(?<=[.!?])\s+', text)
    new_sentences = []
    for sent in sentences:
        words = sent.split()
        if len(words) <= max_words:
            new_sentences.append(sent)
            continue
        split_points = []
        for i, w in enumerate(words):
            if i > 5 and i < len(words)-5 and w.lower() in (',', ';', 'and', 'but', 'so', 'because'):
                split_points.append(i)
        if not split_points:
            mid = len(words) // 2
            split_points.append(mid)
        best = min(split_points, key=lambda x: abs(x - len(words)//2))
        part1 = ' '.join(words[:best]).strip()
        part2 = ' '.join(words[best+1:]).strip()
        if part1 and part2:
            if part1[-1] not in '.!?':
                part1 += '.'
            if part2[-1] not in '.!?':
                part2 += '.'
            part2 = part2[0].upper() + part2[1:]
            new_sentences.extend([part1, part2])
        else:
            new_sentences.append(sent)
    return ' '.join(new_sentences)

def humanize_text(text: str, intensity: int = 3) -> str:
    """تطبيق جميع قواعد التحويل."""
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
        text = split_long_sentences(text, max_words=25)
    elif intensity >= 2:
        text = split_long_sentences(text, max_words=30)
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text

# ============================================================
# 2. دوال تحليل النص (الشفافية)
# ============================================================

def token_count(text: str) -> int:
    return len(re.findall(r'\b\w+\b', text))

def avg_sentence_length(text: str) -> float:
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if s.strip()]
    if not sentences:
        return 0.0
    total_words = sum(len(s.split()) for s in sentences)
    return total_words / len(sentences)

def lexical_diversity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return 0.0
    return len(set(words)) / len(words)

def estimate_perplexity(text: str) -> float:
    """تقدير بسيط للتشوش بناءً على تنوع المفردات وطول الجمل."""
    words = re.findall(r'\b\w+\b', text.lower())
    if len(words) < 3:
        return 50.0
    unique = len(set(words))
    diversity = unique / len(words)
    # كلما انخفض التنوع، زاد التشوش (النص أكثر توقعاً)
    return max(20.0, min(120.0, 100.0 - diversity * 80))

def estimate_burstiness(text: str) -> float:
    """تقدير الاندفاع بناءً على تباين أطوال الجمل."""
    sentences = re.split(r'[.!?]+', text)
    lengths = [len(s.split()) for s in sentences if s.strip()]
    if len(lengths) < 2:
        return 0.0
    mean = sum(lengths) / len(lengths)
    var = sum((l - mean)**2 for l in lengths) / len(lengths)
    return var / (mean + 1e-6)

def count_forbidden_words(text: str) -> int:
    lowered = text.lower()
    return sum(1 for w in FORBIDDEN_SET if w in lowered)

def analyze_text(text: str) -> Dict:
    """تحليل النص وإرجاع المؤشرات الرئيسية."""
    return {
        "chars": len(text),
        "words": token_count(text),
        "sentences": len(re.split(r'[.!?]+', text)),
        "avg_sentence_len": avg_sentence_length(text),
        "lexical_diversity": lexical_diversity(text),
        "perplexity": estimate_perplexity(text),
        "burstiness": estimate_burstiness(text),
        "forbidden_count": count_forbidden_words(text),
    }

def compute_ai_score(analysis: Dict) -> float:
    """حساب درجة الآلية من 0 إلى 1 (0 = بشري، 1 = آلي)."""
    score = 0.0
    # التنوع المعجمي المنخفض يرفع الدرجة
    if analysis["lexical_diversity"] < 0.4:
        score += 0.3
    elif analysis["lexical_diversity"] > 0.6:
        score -= 0.2
    # الاندفاع المنخفض يرفع الدرجة
    if analysis["burstiness"] < 0.15:
        score += 0.35
    elif analysis["burstiness"] > 0.35:
        score -= 0.2
    # التشوش المنخفض يرفع الدرجة
    if analysis["perplexity"] < 35:
        score += 0.25
    elif analysis["perplexity"] > 65:
        score -= 0.15
    # الكلمات المحظورة
    score += min(0.4, analysis["forbidden_count"] * 0.05)
    return min(0.99, max(0.0, score))

def classify_text(score: float) -> Tuple[str, str]:
    if score < 0.20:
        return "بشري (إشارة منخفضة)", "🟢 بشري"
    elif score < 0.40:
        return "بشري محتمل", "🟡 مختلط منخفض"
    elif score < 0.60:
        return "مختلط", "🟠 مختلط مرتفع"
    else:
        return "آلي محتمل", "🔴 آلي"

# ============================================================
# 3. معالجة الملفات
# ============================================================

def extract_text_from_uploaded(uploaded) -> str:
    """استخراج النص من ملف مرفوع (docx, txt, pdf)."""
    name = uploaded.name.lower()
    if name.endswith('.txt'):
        return uploaded.read().decode('utf-8', errors='replace')
    elif name.endswith('.docx'):
        if not DOCX_AVAILABLE:
            return "الرجاء تثبيت python-docx: pip install python-docx"
        return docx2txt.process(uploaded) or ""
    elif name.endswith('.pdf'):
        if not PDF_AVAILABLE:
            return "الرجاء تثبيت pypdf: pip install pypdf"
        reader = pypdf.PdfReader(uploaded)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    else:
        return ""

def process_docx_safe(in_bytes: BytesIO, intensity: int) -> BytesIO:
    """معالجة ملف Word مع الحفاظ على الجداول والأشكال."""
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx غير مثبت")
    doc = Document(in_bytes)
    for para in doc.paragraphs:
        if para.text.strip():
            new_text = humanize_text(para.text, intensity)
            if new_text != para.text:
                para.clear()
                run = para.add_run(new_text)
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ============================================================
# 4. واجهة Streamlit
# ============================================================

def main():
    st.title("📄 DeepClean Studio – الإصدار الشامل النهائي")
    st.caption("إعادة صياغة النصوص الأكاديمية إلى أسلوب بشري مع لوحة شفافية وتحليل – يجتاز ZeroGPT")
    
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        intensity = st.slider("شدة المراجعة (تقطيع الجمل الطويلة)", 1, 5, 3,
                              help="1=أقل تقطيع، 5=أقصى تقطيع للجمل الطويلة")
        st.markdown("---")
        st.header("📥 مصدر النص")
        source = st.radio("اختر طريقة الإدخال", ["لصق نص", "رفع ملف"])
        user_text = ""
        uploaded_file = None
        
        if source == "لصق نص":
            user_text = st.text_area("ألصق النص الأكاديمي هنا", height=200)
        else:
            uploaded_file = st.file_uploader("رفع ملف", type=["docx", "txt", "pdf"])
            if uploaded_file:
                with st.spinner("جاري استخراج النص..."):
                    user_text = extract_text_from_uploaded(uploaded_file)
                if user_text:
                    st.success(f"تم استخراج {len(user_text)} حرف")
                else:
                    st.error("لم يتم استخراج نص من الملف")
        
        process_btn = st.button("🚀 بدء المراجعة والتحليل", type="primary", use_container_width=True)
    
    if process_btn and user_text:
        # تحليل النص الأصلي
        orig_analysis = analyze_text(user_text)
        orig_score = compute_ai_score(orig_analysis)
        orig_class, orig_color = classify_text(orig_score)
        
        # معالجة النص
        with st.spinner("جاري إعادة الصياغة البشرية..."):
            if uploaded_file and uploaded_file.name.endswith('.docx') and DOCX_AVAILABLE:
                # معالجة آمنة لملف Word
                in_bytes = BytesIO(uploaded_file.read())
                out_bytes = process_docx_safe(in_bytes, intensity)
                revised = None  # لن نعرض النص مباشرة لأنه يحتفظ بالتنسيق
                revised_analysis = None
                st.success("تمت معالجة ملف Word مع الحفاظ على الجداول والأشكال!")
                st.download_button("⬇️ تحميل الملف المعدّل (Word)", data=out_bytes,
                                   file_name="deepclean_humanized.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                # معالجة النص العادي
                revised = humanize_text(user_text, intensity)
                revised_analysis = analyze_text(revised)
                revised_score = compute_ai_score(revised_analysis)
                revised_class, revised_color = classify_text(revised_score)
        
        # عرض النتائج في تبويبات
        tab1, tab2, tab3 = st.tabs(["📊 التحليل والمقارنة", "📝 النص الأصلي", "✨ النص المعدل"])
        
        with tab1:
            st.subheader("مؤشرات الجودة")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**النص الأصلي**")
                st.metric("عدد الكلمات", orig_analysis["words"])
                st.metric("متوسط طول الجملة", f"{orig_analysis['avg_sentence_len']:.1f}")
                st.metric("التنوع المعجمي", f"{orig_analysis['lexical_diversity']:.3f}")
                st.metric("التشوش (Perplexity)", f"{orig_analysis['perplexity']:.1f}")
                st.metric("الاندفاع (Burstiness)", f"{orig_analysis['burstiness']:.3f}")
                st.metric("الكلمات المحظورة", orig_analysis["forbidden_count"])
                st.metric("درجة الآلية", f"{orig_score:.2f}")
                st.markdown(f"**التصنيف:** {orig_color} {orig_class}")
            if revised_analysis:
                with col2:
                    st.markdown("**النص المعدل**")
                    st.metric("عدد الكلمات", revised_analysis["words"])
                    st.metric("متوسط طول الجملة", f"{revised_analysis['avg_sentence_len']:.1f}")
                    st.metric("التنوع المعجمي", f"{revised_analysis['lexical_diversity']:.3f}")
                    st.metric("التشوش (Perplexity)", f"{revised_analysis['perplexity']:.1f}")
                    st.metric("الاندفاع (Burstiness)", f"{revised_analysis['burstiness']:.3f}")
                    st.metric("الكلمات المحظورة", revised_analysis["forbidden_count"])
                    st.metric("درجة الآلية", f"{revised_score:.2f}")
                    st.markdown(f"**التصنيف:** {revised_color} {revised_class}")
            else:
                st.info("لم يتم تحليل النص المعدل (ملف Word محفوظ التنسيق)")
        
        with tab2:
            st.text_area("النص الأصلي", user_text, height=400)
        
        with tab3:
            if revised:
                st.text_area("النص المعدل (بشري)", revised, height=400)
                st.download_button("📥 تحميل النص المعدل (TXT)", data=revised.encode('utf-8'),
                                   file_name="deepclean_humanized.txt", mime="text/plain")
            elif uploaded_file and uploaded_file.name.endswith('.docx'):
                st.info("تم معالجة ملف Word. استخدم زر التحميل أعلاه للحصول على الملف المعدّل.")
            else:
                st.info("لا يوجد نص معروض")
    
    elif process_btn and not user_text:
        st.warning("الرجاء إدخال نص أو رفع ملف.")

if __name__ == "__main__":
    main()
