#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - الإصدار النهائي (معالجة آمنة لـ Word واللصق المباشر)
يعالج النصوص الأكاديمية المستخرجة من ملفات Word أو النص المباشر،
ويطبق قواعد إعادة الصياغة البشرية، ويعرض النص المعدل لنسخه ولصقه يدويًا.
"""

import re
import streamlit as st
from io import BytesIO
from typing import Dict

# مكتبات استخراج النص من Word
try:
    import docx2txt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - Word & Paste", layout="wide")

# ============================================================
# 1. قواعد إعادة الصياغة البشرية (موثقة من التجارب السابقة)
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

def split_long_sentences(text: str, max_words: int = 28) -> str:
    """تقطيع الجمل الطويلة إلى جملتين، مع الحفاظ على النحو السليم."""
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
    new_sentences = []
    for sent in sentences:
        words = sent.split()
        if len(words) <= max_words:
            new_sentences.append(sent)
            continue
        # البحث عن فاصلة أو حرف عطف
        split_pos = -1
        for i, w in enumerate(words):
            if i > 6 and i < len(words)-4 and w.lower() in (',', 'and', 'but', 'so', 'because', 'while', 'whereas'):
                split_pos = i
                break
        if split_pos > 0:
            first = ' '.join(words[:split_pos]).strip()
            second = ' '.join(words[split_pos+1:]).strip()
            if first and second:
                if first[-1] not in '.!?':
                    first += '.'
                if second[-1] not in '.!?':
                    second += '.'
                second = second[0].upper() + second[1:]
                new_sentences.extend([first, second])
            else:
                new_sentences.append(sent)
        else:
            mid = len(words) // 2
            first = ' '.join(words[:mid]).strip()
            second = ' '.join(words[mid:]).strip()
            if first and second:
                if first[-1] not in '.!?':
                    first += '.'
                if second[-1] not in '.!?':
                    second += '.'
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
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text.strip()

# ============================================================
# 2. دوال التحليل
# ============================================================

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
    return max(20.0, min(120.0, 100.0 - diversity * 80))

def estimate_burstiness(text: str) -> float:
    lengths = [len(s.split()) for s in re.split(r'[.!?]+', text) if s.strip()]
    if len(lengths) < 2:
        return 0.0
    mean = sum(lengths) / len(lengths)
    var = sum((l - mean)**2 for l in lengths) / len(lengths)
    return var / (mean + 1e-6)

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
        "burstiness": estimate_burstiness(text),
        "forbidden": count_forbidden_words(text),
    }

def ai_score(analysis: Dict) -> float:
    score = 0.0
    if analysis["lexical_diversity"] < 0.4:
        score += 0.3
    elif analysis["lexical_diversity"] > 0.65:
        score -= 0.2
    if analysis["burstiness"] < 0.15:
        score += 0.35
    elif analysis["burstiness"] > 0.4:
        score -= 0.2
    if analysis["perplexity"] < 35:
        score += 0.25
    elif analysis["perplexity"] > 65:
        score -= 0.15
    score += min(0.4, analysis["forbidden"] * 0.05)
    return min(0.99, max(0.0, score))

def classify_score(score: float) -> str:
    if score < 0.20:
        return "🟢 بشري (إشارة منخفضة)"
    elif score < 0.40:
        return "🟡 بشري محتمل"
    elif score < 0.60:
        return "🟠 مختلط"
    else:
        return "🔴 آلي محتمل"

# ============================================================
# 3. استخراج النص من ملف Word
# ============================================================

def extract_text_from_word(file_bytes: BytesIO) -> str:
    if not DOCX_AVAILABLE:
        return "الرجاء تثبيت docx2txt: pip install docx2txt"
    try:
        text = docx2txt.process(file_bytes) or ""
        return text
    except Exception as e:
        return f"خطأ في قراءة الملف: {e}"

# ============================================================
# 4. واجهة المستخدم
# ============================================================

def main():
    st.title("📄 DeepClean Studio – معالجة Word واللصق المباشر")
    st.caption("يعالج النصوص الأكاديمية المستخرجة من ملفات Word أو النص المباشر، ويحولها إلى أسلوب بشري مع تحليل كامل.")
    st.warning("⚠️ للحفاظ على الجداول والأشكال والمعادلات في ملف Word الأصلي: انسخ النص المعدل والصقه يدويًا في المستند الأصلي (استبدل الفقرات النصية فقط).")

    with st.sidebar:
        st.header("⚙️ الإعدادات")
        intensity = st.slider("شدة المراجعة", 1, 5, 3,
                              help="كلما زادت الشدة، زاد تقطيع الجمل الطويلة واستبدال الكلمات")
        st.markdown("---")
        st.header("📥 المصدر")
        source = st.radio("اختر طريقة الإدخال", ["لصق النص مباشرة", "رفع ملف Word"])
        user_text = ""

        if source == "لصق النص مباشرة":
            user_text = st.text_area("ألصق النص الأكاديمي هنا (مقدمة، خاتمة، مناقشة...)", height=250)
        else:  # رفع ملف Word
            uploaded = st.file_uploader("اختر ملف Word (.docx)", type=["docx"])
            if uploaded:
                with st.spinner("جاري استخراج النص من Word..."):
                    user_text = extract_text_from_word(BytesIO(uploaded.read()))
                if user_text and not user_text.startswith("خطأ"):
                    st.success(f"تم استخراج {len(user_text)} حرف")
                elif user_text.startswith("خطأ"):
                    st.error(user_text)
                else:
                    st.warning("لم يتم استخراج نص من الملف")

        process = st.button("🚀 بدء المراجعة والتحليل", type="primary", use_container_width=True)

    if process and user_text and not user_text.startswith("خطأ"):
        with st.spinner("جاري المعالجة..."):
            orig_analysis = analyze_text(user_text)
            orig_score = ai_score(orig_analysis)
            revised = humanize_text(user_text, intensity)
            rev_analysis = analyze_text(revised)
            rev_score = ai_score(rev_analysis)

        # عرض المقارنة
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📊 النص الأصلي")
            st.metric("عدد الكلمات", orig_analysis["words"])
            st.metric("متوسط طول الجملة", f"{orig_analysis['avg_sentence_len']:.1f}")
            st.metric("التنوع المعجمي", f"{orig_analysis['lexical_diversity']:.3f}")
            st.metric("التشوش (perplexity)", f"{orig_analysis['perplexity']:.1f}")
            st.metric("الاندفاع (burstiness)", f"{orig_analysis['burstiness']:.3f}")
            st.metric("كلمات محظورة", orig_analysis["forbidden"])
            st.metric("درجة الآلية", f"{orig_score:.2f}")
            st.write(f"**التصنيف:** {classify_score(orig_score)}")

        with col2:
            st.subheader("✨ النص المعدل (بشري)")
            st.metric("عدد الكلمات", rev_analysis["words"])
            st.metric("متوسط طول الجملة", f"{rev_analysis['avg_sentence_len']:.1f}")
            st.metric("التنوع المعجمي", f"{rev_analysis['lexical_diversity']:.3f}")
            st.metric("التشوش", f"{rev_analysis['perplexity']:.1f}")
            st.metric("الاندفاع", f"{rev_analysis['burstiness']:.3f}")
            st.metric("كلمات محظورة", rev_analysis["forbidden"])
            st.metric("درجة الآلية", f"{rev_score:.2f}")
            st.write(f"**التصنيف:** {classify_score(rev_score)}")

        # عرض النصوص المعدلة
        tab1, tab2 = st.tabs(["📝 النص الأصلي", "✨ النص المعدل (انسخه هنا)"])
        with tab1:
            st.text_area("", user_text, height=400, key="orig_text")
        with tab2:
            st.text_area("", revised, height=400, key="rev_text")
            col_a, col_b = st.columns(2)
            with col_a:
                st.download_button("📥 تحميل النص المعدل (TXT)", data=revised.encode('utf-8'),
                                   file_name="deepclean_humanized.txt", mime="text/plain")
            with col_b:
                st.button("📋 نسخ إلى الحافظة (يدويًا)", disabled=True,
                          help="حدد النص أعلاه واضغط Ctrl+C")
            st.info("💡 **كيفية الاستخدام الآمن في Word:**\n"
                    "1. افتح ملف Word الأصلي.\n"
                    "2. حدد الفقرات النصية التي تريد استبدالها (مثل المقدمة).\n"
                    "3. انسخ النص المعدل من المربع أعلاه.\n"
                    "4. الصق النص في Word (استخدم 'لصق مع الاحتفاظ بالنص فقط').\n"
                    "5. تأكد من بقاء الجداول والأشكال والمعادلات كما هي.\n\n"
                    "✅ هذه الطريقة تضمن عدم تلف المستند الأصلي.")

    elif process:
        if not user_text:
            st.warning("الرجاء إدخال نص أو رفع ملف Word.")
        elif user_text.startswith("خطأ"):
            st.error(user_text)

if __name__ == "__main__":
    main()
