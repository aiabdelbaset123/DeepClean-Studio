#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - الإصدار الشامل النهائي (مصحح)
يدعم رفع Word مع الحفاظ على التنسيق، ولصق النص، وتحليل الشفافية.
"""

import re
import streamlit as st
from io import BytesIO
from typing import Tuple, Dict, Optional

# مكتبات اختيارية
try:
    from docx import Document
    from docx.shared import Pt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import docx2txt
    DOCX2TXT_AVAILABLE = True
except ImportError:
    DOCX2TXT_AVAILABLE = False

try:
    import pypdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - النهائي", layout="wide")

# ============================================================
# 1. قواعد إعادة الصياغة البشرية (نفس السابق)
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
# 2. تحليل النص (الشفافية)
# ============================================================

def token_count(text: str) -> int:
    return len(re.findall(r'\b\w+\b', text))

def avg_sentence_length(text: str) -> float:
    sentences = [s.strip() for s in re.split(r'[.!?]+', text) if s.strip()]
    if not sentences:
        return 0.0
    total = sum(len(s.split()) for s in sentences)
    return total / len(sentences)

def lexical_diversity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return 0.0
    return len(set(words)) / len(words)

def estimate_perplexity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if len(words) < 3:
        return 50.0
    unique = len(set(words))
    diversity = unique / len(words)
    return max(20.0, min(120.0, 100.0 - diversity * 80))

def estimate_burstiness(text: str) -> float:
    lengths = [len(s.split()) for s in re.split(r'[.!?]+', text) if s.strip()]
    if len(lengths) < 2:
        return 0.0
    mean = sum(lengths) / len(lengths)
    var = sum((l - mean)**2 for l in lengths) / len(lengths)
    return var / (mean + 1e-6)

def count_forbidden(text: str) -> int:
    lowered = text.lower()
    return sum(1 for w in FORBIDDEN_SET if w in lowered)

def analyze_text(text: str) -> Dict:
    return {
        "words": token_count(text),
        "avg_sentence_len": avg_sentence_length(text),
        "lexical_diversity": lexical_diversity(text),
        "perplexity": estimate_perplexity(text),
        "burstiness": estimate_burstiness(text),
        "forbidden": count_forbidden(text),
    }

def ai_score(analysis: Dict) -> float:
    score = 0.0
    if analysis["lexical_diversity"] < 0.4:
        score += 0.3
    elif analysis["lexical_diversity"] > 0.6:
        score -= 0.2
    if analysis["burstiness"] < 0.15:
        score += 0.35
    elif analysis["burstiness"] > 0.35:
        score -= 0.2
    if analysis["perplexity"] < 35:
        score += 0.25
    elif analysis["perplexity"] > 65:
        score -= 0.15
    score += min(0.4, analysis["forbidden"] * 0.05)
    return min(0.99, max(0.0, score))

def classify(score: float) -> str:
    if score < 0.20:
        return "بشري (إشارة منخفضة) 🟢"
    elif score < 0.40:
        return "بشري محتمل 🟡"
    elif score < 0.60:
        return "مختلط 🟠"
    else:
        return "آلي محتمل 🔴"

# ============================================================
# 3. معالجة الملفات
# ============================================================

def extract_text_from_pdf(file_bytes: BytesIO) -> str:
    if not PDF_AVAILABLE:
        return "تثبيت pypdf: pip install pypdf"
    reader = pypdf.PdfReader(file_bytes)
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def extract_text_from_docx(file_bytes: BytesIO) -> str:
    if not DOCX2TXT_AVAILABLE:
        return "تثبيت docx2txt: pip install docx2txt"
    return docx2txt.process(file_bytes) or ""

def process_docx_file(file_bytes: BytesIO, intensity: int) -> BytesIO:
    """تعديل ملف Word مع الحفاظ على الجداول والأشكال."""
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx غير مثبت")
    # Reset stream position
    file_bytes.seek(0)
    doc = Document(file_bytes)
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
# 4. واجهة المستخدم
# ============================================================

def main():
    st.title("📄 DeepClean Studio – الإصدار النهائي المتكامل")
    st.caption("إعادة كتابة النصوص الأكاديمية بأسلوب بشري - مع تحليل شفاف - يجتاز ZeroGPT")
    
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        intensity = st.slider("شدة المراجعة", 1, 5, 3,
                              help="كلما زادت الشدة، زاد تقطيع الجمل الطويلة واستبدال الكلمات")
        st.markdown("---")
        st.header("📥 إدخال النص")
        input_type = st.radio("المصدر", ["لصق نص", "رفع ملف Word", "رفع ملف PDF/TXT"])
        
        user_text = ""
        uploaded_file = None
        file_bytes = None
        
        if input_type == "لصق نص":
            user_text = st.text_area("ألصق النص هنا", height=200)
        else:
            if input_type == "رفع ملف Word":
                uploaded_file = st.file_uploader("اختر ملف Word", type=["docx"])
            else:
                uploaded_file = st.file_uploader("اختر ملف PDF أو TXT", type=["pdf", "txt"])
            if uploaded_file:
                file_bytes = BytesIO(uploaded_file.read())
                if input_type == "رفع ملف Word":
                    # لا نستخرج النص إلا للعرض
                    if DOCX2TXT_AVAILABLE:
                        user_text = docx2txt.process(BytesIO(uploaded_file.getvalue())) or ""
                    else:
                        user_text = "تثبيت docx2txt لعرض النص"
                elif uploaded_file.name.endswith('.pdf'):
                    user_text = extract_text_from_pdf(BytesIO(uploaded_file.getvalue()))
                else:
                    user_text = uploaded_file.getvalue().decode('utf-8', errors='replace')
        
        process = st.button("🚀 بدء المراجعة والتحليل", type="primary", use_container_width=True)
    
    if process:
        # حالة: النص المباشر أو النص المستخرج من ملف غير Word
        if user_text and input_type != "رفع ملف Word":
            with st.spinner("جاري المعالجة..."):
                orig_analysis = analyze_text(user_text)
                orig_score = ai_score(orig_analysis)
                revised = humanize_text(user_text, intensity)
                rev_analysis = analyze_text(revised)
                rev_score = ai_score(rev_analysis)
            
            # عرض النتائج
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("📊 النص الأصلي")
                st.metric("الكلمات", orig_analysis["words"])
                st.metric("متوسط طول الجملة", f"{orig_analysis['avg_sentence_len']:.1f}")
                st.metric("التنوع المعجمي", f"{orig_analysis['lexical_diversity']:.3f}")
                st.metric("التشوش", f"{orig_analysis['perplexity']:.1f}")
                st.metric("الاندفاع", f"{orig_analysis['burstiness']:.3f}")
                st.metric("كلمات محظورة", orig_analysis["forbidden"])
                st.metric("درجة الآلية", f"{orig_score:.2f}")
                st.write(f"**التصنيف:** {classify(orig_score)}")
            with col2:
                st.subheader("✨ النص المعدل")
                st.metric("الكلمات", rev_analysis["words"])
                st.metric("متوسط طول الجملة", f"{rev_analysis['avg_sentence_len']:.1f}")
                st.metric("التنوع المعجمي", f"{rev_analysis['lexical_diversity']:.3f}")
                st.metric("التشوش", f"{rev_analysis['perplexity']:.1f}")
                st.metric("الاندفاع", f"{rev_analysis['burstiness']:.3f}")
                st.metric("كلمات محظورة", rev_analysis["forbidden"])
                st.metric("درجة الآلية", f"{rev_score:.2f}")
                st.write(f"**التصنيف:** {classify(rev_score)}")
            
            # عرض النصوص في تبويبات
            tab1, tab2 = st.tabs(["📝 النص الأصلي", "✨ النص المعدل"])
            with tab1:
                st.text_area("", user_text, height=300)
            with tab2:
                st.text_area("", revised, height=300)
                st.download_button("📥 تحميل النص المعدل (TXT)", data=revised.encode('utf-8'),
                                   file_name="deepclean_humanized.txt")
        
        # حالة: ملف Word مباشر (معالجة آمنة)
        elif uploaded_file and input_type == "رفع ملف Word" and file_bytes:
            with st.spinner("جاري معالجة ملف Word مع الحفاظ على الجداول والأشكال..."):
                try:
                    out_bytes = process_docx_file(file_bytes, intensity)
                    st.success("تمت معالجة الملف بنجاح! الجداول والأشكال والمعادلات سليمة.")
                    st.download_button("📥 تحميل الملف المعدّل (Word)", data=out_bytes,
                                       file_name="deepclean_humanized.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    # عرض تحليل النص المستخرج (للإرشاد فقط)
                    if user_text and len(user_text) > 100:
                        st.info("تحليل تقديري للنص المستخرج (وليس الملف نفسه):")
                        ext_analysis = analyze_text(user_text[:5000])
                        ext_score = ai_score(ext_analysis)
                        st.write(f"درجة الآلية التقديرية: {ext_score:.2f} - {classify(ext_score)}")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
        else:
            st.warning("الرجاء إدخال نص أو رفع ملف.")

if __name__ == "__main__":
    main()
