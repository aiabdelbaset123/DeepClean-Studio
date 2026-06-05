#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Humanize Engine (Lightweight)
استخدام القواعد والإحصاءات البسيطة بدلاً من النماذج الثقيلة.
يعمل بدون أخطاء استيراد على Streamlit Cloud.
"""

import re
import random
import math
from io import BytesIO
from typing import List, Dict, Optional

import streamlit as st
import numpy as np
from collections import Counter

# استيراد المكتبات الخفيفة لرفع الملفات
try:
    import docx2txt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pypdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - Humanize Engine", layout="wide")

# ----------------------------------------------------------------------
# قاعدة المرادفات (مضمنة، لا تحتاج ملف خارجي)
# ----------------------------------------------------------------------
SYNONYMS = {
    "medical": {
        "show": ["demonstrate", "indicate", "reveal", "suggest", "exhibit"],
        "important": ["critical", "salient", "noteworthy", "paramount", "essential"],
        "cause": ["induce", "elicit", "provoke", "trigger", "initiate"],
        "effect": ["impact", "consequence", "outcome", "sequela", "repercussion"],
        "increase": ["elevate", "augment", "raise", "boost", "escalate"],
        "decrease": ["reduce", "diminish", "lower", "attenuate", "curtail"],
        "patients": ["subjects", "cohort", "individuals", "participants"],
        "data": ["findings", "evidence", "observations", "results"],
        "analysis": ["assessment", "evaluation", "appraisal", "examination"],
        "correlation": ["association", "link", "relationship", "connection"],
    },
    "engineering": {
        "show": ["demonstrate", "illustrate", "exhibit", "reveal", "display"],
        "important": ["crucial", "vital", "essential", "key", "critical"],
        "change": ["modify", "alter", "adjust", "transform", "reshape"],
        "use": ["employ", "utilize", "apply", "deploy", "leverage"],
        "increase": ["boost", "enhance", "raise", "amplify", "escalate"],
        "decrease": ["reduce", "lower", "attenuate", "dampen", "diminish"],
        "performance": ["efficiency", "output", "throughput", "capability"],
        "system": ["assembly", "setup", "configuration", "architecture"],
        "process": ["procedure", "methodology", "workflow", "routine"],
        "design": ["layout", "topology", "configuration", "blueprint"],
    },
    "humanities": {
        "show": ["demonstrate", "reveal", "expose", "lay bare", "uncover"],
        "important": ["significant", "consequential", "notable", "weighty", "momentous"],
        "argue": ["contend", "assert", "maintain", "posit", "allege"],
        "believe": ["hold", "maintain", "submit", "allege", "conjecture"],
        "influence": ["shape", "mold", "affect", "impinge on", "determine"],
        "change": ["transform", "reshape", "alter", "shift", "metamorphose"],
        "social": ["societal", "communal", "collective", "interpersonal"],
        "culture": ["civilization", "society", "milieu", "ethos"],
        "meaning": ["significance", "import", "sense", "purport"],
        "context": ["setting", "background", "frame", "circumstance"],
    },
    "general": {
        "show": ["demonstrate", "indicate", "suggest", "reveal", "illustrate"],
        "important": ["significant", "notable", "considerable", "substantial", "major"],
        "change": ["modify", "alter", "transform", "adjust", "vary"],
        "use": ["employ", "utilize", "apply", "deploy", "operate"],
        "increase": ["raise", "boost", "augment", "elevate", "amplify"],
        "decrease": ["reduce", "lower", "diminish", "curtail", "lessen"],
        "data": ["findings", "results", "observations", "evidence"],
        "analysis": ["examination", "study", "investigation", "evaluation"],
        "method": ["approach", "technique", "procedure", "methodology"],
        "result": ["outcome", "finding", "consequence", "product"],
    }
}

# ----------------------------------------------------------------------
# دوال مساعدة لتحليل النص (خفيفة)
# ----------------------------------------------------------------------
def split_sentences(text: str) -> List[str]:
    """تقسيم النص إلى جمل باستخدام علامات الترقيم الأساسية."""
    text = re.sub(r'\s+', ' ', text)
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
    return [s.strip() for s in sentences if s.strip()]

def estimate_perplexity(text: str) -> float:
    """تقدير سريع للتشوش بناءً على تنوع المفردات وطول الجمل."""
    words = re.findall(r'\b\w+\b', text.lower())
    if len(words) < 3:
        return 50.0
    unique = len(set(words))
    diversity = unique / len(words)
    return max(20.0, min(120.0, 100.0 - diversity * 70))

def burstiness_score(text: str) -> float:
    """تقدير الاندفاع (تباين أطوال الجمل)."""
    lens = [len(s.split()) for s in split_sentences(text)]
    if len(lens) < 2:
        return 0.0
    mean = np.mean(lens)
    var = np.var(lens)
    return var / (mean + 1e-6)

def lexical_diversity(text: str) -> float:
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return 0.0
    return len(set(words)) / len(words)

def count_forbidden_words(text: str) -> int:
    """عدد الكلمات المحظورة (علامات الذكاء الاصطناعي)."""
    forbidden = {
        "additionally", "moreover", "furthermore", "consequently", "hence",
        "crucial", "pivotal", "vital", "significant", "profound", "robust",
        "comprehensive", "delve", "showcase", "underscore", "highlight",
        "resonate", "align", "garner", "tapestry", "testament", "landscape",
        "intricate", "multifaceted", "constitute", "trajectories", "pronounced",
        "routinely", "impose", "exceeding", "constituting", "cumulative", "uniquely"
    }
    lowered = text.lower()
    return sum(1 for w in forbidden if w in lowered)

def display_metrics(text: str, label: str) -> None:
    """عرض مؤشرات الجودة في واجهة المستخدم."""
    words = len(text.split())
    sent_len = np.mean([len(s.split()) for s in split_sentences(text)]) if split_sentences(text) else 0
    lex_div = lexical_diversity(text)
    pp = estimate_perplexity(text)
    burst = burstiness_score(text)
    forbidden = count_forbidden_words(text)
    st.metric(f"{label} - Word count", words)
    st.caption(f"Avg sentence length: {sent_len:.1f} | Lexical diversity: {lex_div:.3f}")
    st.caption(f"Perplexity (est): {pp:.1f} | Burstiness: {burst:.3f} | Forbidden words: {forbidden}")

# ----------------------------------------------------------------------
# محرك التأمين البشري (خفيف)
# ----------------------------------------------------------------------
class HumanizeEngine:
    def __init__(self, text: str, domain: str, strength: int):
        self.original = text
        self.domain = domain
        self.strength = min(5, max(1, strength))
        self.synonyms = SYNONYMS.get(domain, SYNONYMS["general"])
        random.seed(42)   # للتكرار

    def _replace_forbidden_words(self, text: str) -> str:
        """استبدال الكلمات المحظورة بمرادفات بشرية."""
        for old, syn_list in self.synonyms.items():
            if old in text.lower():
                # 70% من المرات نستبدل
                if random.random() < 0.7:
                    new = random.choice(syn_list)
                    text = re.compile(rf'\b{re.escape(old)}\b', re.IGNORECASE).sub(new, text)
        return text

    def _split_long_sentences(self, sentences: List[str]) -> List[str]:
        """تقطيع الجمل الطويلة جداً (>35 كلمة) إلى جملتين."""
        new_sentences = []
        for sent in sentences:
            words = sent.split()
            if len(words) > 35 and self.strength >= 3:
                # نقطة التقطيع: فاصلة أو كلمة ربط
                split_pos = -1
                for i, w in enumerate(words):
                    if w.lower() in (',', 'and', 'but', 'so', 'because') and 10 < i < len(words)-10:
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
                    mid = len(words)//2
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
            else:
                new_sentences.append(sent)
        return new_sentences

    def _vary_sentence_length(self, sentences: List[str]) -> List[str]:
        """ضمان وجود جمل قصيرة جداً (<8 كلمات) وأخرى طويلة (>25)."""
        if len(sentences) < 2:
            return sentences
        short_count = sum(1 for s in sentences if len(s.split()) < 8)
        long_count = sum(1 for s in sentences if len(s.split()) > 25)
        # إضافة جملة قصيرة إذا لم توجد
        if short_count == 0 and self.strength >= 2:
            # تقطيع أطول جملة لإنتاج جزء قصير
            longest = max(sentences, key=lambda x: len(x.split()))
            words = longest.split()
            if len(words) > 8:
                short_part = ' '.join(words[:3]) + '.'
                rest = ' '.join(words[3:])
                sentences.remove(longest)
                sentences.append(rest)
                sentences.append(short_part)
        # إضافة جملة طويلة إذا لم توجد (بدمج جملتين قصيرتين)
        if long_count == 0 and len(sentences) > 2 and self.strength >= 2:
            for i in range(len(sentences)-1):
                if len(sentences[i].split()) < 12 and len(sentences[i+1].split()) < 12:
                    merged = sentences[i] + ' ' + sentences[i+1][0].lower() + sentences[i+1][1:]
                    sentences[i] = merged
                    del sentences[i+1]
                    break
        return sentences

    def _add_human_touches(self, text: str) -> str:
        """إضافة لمسات بشرية خفيفة: بدايات عامية، تحفظات، أخطاء نادرة."""
        sentences = split_sentences(text)
        for i, sent in enumerate(sentences):
            # إضافة "So" أو "Well" لبعض الجمل الأولى (مرة واحدة)
            if i == 0 and random.random() < 0.2:
                starters = ["So, ", "Well, ", "Look, ", "I mean, "]
                sentences[i] = random.choice(starters) + sent[0].lower() + sent[1:]
            # إضافة "right?" في النهاية (نادراً)
            if self.strength >= 4 and random.random() < 0.05:
                sentences[i] = sent.rstrip('.!?') + ', right?'
        return ' '.join(sentences)

    def _protect_citations(self, original: str, modified: str) -> str:
        """استعادة الاستشهادات والأرقام كما كانت في النص الأصلي."""
        citations = re.findall(r'\[\d+(?:[-,;]\s*\d+)*\]', original)
        numbers = re.findall(r'\b\d+(?:\.\d+)?\s?(?:%|°C|GW|kWh|W/m²|km|m|kg|s)?\b', original)
        # إزالة أي شيء يشبهها من النص المعدل
        for c in citations:
            modified = modified.replace(c, '')
        for n in numbers:
            modified = modified.replace(n, '')
        # إضافتها مرة أخرى في نهاية الجمل المناسبة (الحفاظ على الترتيب)
        # نضيفها جميعاً في نهاية النص
        modified = modified.rstrip('.!?') + ' ' + ' '.join(citations + numbers) + '.'
        return modified

    def humanize(self) -> str:
        """تنفيذ جميع عمليات التحويل."""
        text = self.original
        # 1. استبدال الكلمات المحظورة
        text = self._replace_forbidden_words(text)
        # 2. تقسيم إلى جمل
        sentences = split_sentences(text)
        # 3. تقطيع الجمل الطويلة
        sentences = self._split_long_sentences(sentences)
        # 4. تنويع أطوال الجمل
        sentences = self._vary_sentence_length(sentences)
        # 5. إضافة لمسات بشرية
        text = ' '.join(sentences)
        text = self._add_human_touches(text)
        # 6. حماية الاستشهادات والأرقام
        text = self._protect_citations(self.original, text)
        # 7. تنظيف نهائي
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s+([.,;:!?])', r'\1', text)
        return text.strip()

# ----------------------------------------------------------------------
# واجهة المستخدم
# ----------------------------------------------------------------------
def main():
    st.title("🧬 DeepClean Studio – Humanize Engine (Lightweight)")
    st.caption("يحول النصوص الأكاديمية إلى أسلوب بشري يجتاز كاشفات الذكاء الاصطناعي (ZeroGPT، GPTZero، إلخ).")
    st.sidebar.header("⚙️ الإعدادات")

    # مصدر النص
    input_type = st.sidebar.radio("مصدر النص", ["لصق نص", "رفع ملف"])
    user_text = ""
    if input_type == "لصق نص":
        user_text = st.sidebar.text_area("ألصق النص هنا", height=200)
    else:
        uploaded = st.sidebar.file_uploader("اختر ملف (txt, docx, pdf)", type=["txt", "docx", "pdf"])
        if uploaded:
            ext = uploaded.name.split('.')[-1].lower()
            if ext == "txt":
                user_text = uploaded.read().decode("utf-8")
            elif ext == "docx" and DOCX_AVAILABLE:
                user_text = docx2txt.process(uploaded)
            elif ext == "pdf" and PDF_AVAILABLE:
                reader = pypdf.PdfReader(uploaded)
                user_text = "\n".join(page.extract_text() or "" for page in reader.pages)
            else:
                st.sidebar.error("تنسيق غير مدعوم أو مكتبة مفقودة")

    domain = st.sidebar.selectbox("المجال الأكاديمي", ["general", "medical", "engineering", "humanities"])
    strength = st.sidebar.slider("قوة التحويل", 1, 5, 3, help="1=تحفظ، 5=ابتكاري")
    process = st.sidebar.button("🛡️ بدء التأمين البشري", type="primary")

    if process and user_text:
        with st.spinner("جاري تحويل النص..."):
            engine = HumanizeEngine(user_text, domain, strength)
            humanized = engine.humanize()
            st.session_state['humanized'] = humanized
            st.session_state['original'] = user_text

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📄 النص الأصلي")
            st.text_area("", user_text, height=400, key="orig")
            display_metrics(user_text, "الأصلي")
        with col2:
            st.subheader("🧬 النص المعدل (بشري)")
            st.text_area("", humanized, height=400, key="human")
            display_metrics(humanized, "المعدل")
            st.download_button("⬇️ تحميل النص المعدل", data=humanized.encode('utf-8'),
                               file_name="deepclean_humanized.txt", mime="text/plain")
    elif process:
        st.sidebar.warning("الرجاء إدخال نص أو رفع ملف.")

if __name__ == "__main__":
    main()
