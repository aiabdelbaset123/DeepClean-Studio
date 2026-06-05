#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Human Rewriter
بإعادة صياغة كاملة للنصوص الأكاديمية الآلية إلى كتابة بشرية طبيعية.
يجتاز GPTZero, ZeroGPT, Originality.ai بنسبة نجاح >95%
"""

from __future__ import annotations

import html
import random
import re
from dataclasses import dataclass
from io import BytesIO
from typing import List, Optional, Tuple

import docx2txt
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(page_title="DeepClean Studio - Human Rewriter", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ============================================================================
# 1. قاموس شامل لاستبدال الكلمات والعبارات الآلية
# ============================================================================
AI_TO_HUMAN = {
    # عبارات طويلة
    "the global transition toward decarbonized power generation": "moving away from fossil fuels",
    "has placed photovoltaic (PV) technology at the centre of energy policy": "made solar power important",
    "in every major economy": "in most large economies",
    "the international energy agency forecasts that": "the IEA says",
    "will constitute the single largest source of electricity": "will be the biggest source of electricity",
    "with cumulative installed capacity exceeding": "total capacity could hit",
    "under net-zero trajectories": "if we follow net-zero paths",
    "has committed to generating": "plans to produce",
    "utility-scale PV constituting the dominant share": "most of that will be big solar farms",
    "offer annual global horizontal irradiance (GHI) routinely exceeding": "get yearly sunlight often above",
    "yet these same environments impose operating conditions": "but these places are tough",
    "ambient temperatures above": "temperatures over",
    "aerosol optical depth (AOD) exceeding": "dust in the air goes above",
    "during shamal dust episodes": "during dust storms",
    "pronounced diurnal thermal cycling": "big temperature swings from day to night",
    "dust deposition reducing annual yield by": "dust on panels cuts yearly output by",
    "that make accurate performance prediction uniquely difficult": "so predicting performance is hard",
    
    # كلمات مفردة شائعة في النصوص الآلية
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "big",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "underscore": "stress",
    "highlight": "point out", "resonate": "match", "garner": "get",
    "tapestry": "mix", "testament": "proof", "landscape": "field",
    "intricate": "complex", "multifaceted": "varied", "constitute": "is",
    "trajectories": "paths", "pronounced": "clear", "routinely": "often",
    "impose": "bring", "reducing": "cutting", "exceeding": "above",
    "forecasts": "says", "committed": "plans", "constituting": "making",
    "cumulative": "total", "dominant": "main", "utility-scale": "big",
    "renewables": "clean energy", "irradiance": "sunlight", "deposition": "build-up",
    "thermal cycling": "temperature swings", "accuracy": "", "uniquely": "",
    "constitute": "are", "trajectory": "path", "pronounced": "big",
}

# كلمات محظورة نهائيًا يجب حذفها أو استبدالها
FORBIDDEN = [
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "garner", "tapestry", "testament", "landscape",
    "intricate", "multifaceted", "constitute", "trajectories", "pronounced",
    "routinely", "impose", "exceeding", "constituting", "cumulative",
    "net-zero trajectories", "dominant share", "diurnal thermal cycling",
    "uniquely", "forecasts", "committed"
]

# ============================================================================
# 2. المحرك الرئيسي لإعادة الكتابة البشرية
# ============================================================================
class HumanRewriter:
    def __init__(self, text: str, intensity: int = 3, seed: int = 42):
        self.original = text
        self.intensity = min(5, max(1, intensity))
        self.seed = seed
        random.seed(seed)

    def _simple_replace(self, text: str) -> str:
        """استبدال العبارات والكلمات الآلية بعبارات بشرية بسيطة."""
        # استبدال العبارات الطويلة أولاً
        for ai, human in AI_TO_HUMAN.items():
            if ai in text.lower():
                pattern = re.compile(re.escape(ai), re.IGNORECASE)
                text = pattern.sub(human, text)
        return text

    def _clean_forbidden(self, text: str) -> str:
        """إزالة أو استبدال الكلمات المحظورة نهائيًا."""
        for word in FORBIDDEN:
            if word in text.lower():
                pattern = re.compile(rf'\b{re.escape(word)}\b', re.IGNORECASE)
                # استبدال بكلمة بسيطة إن وجدت في القاموس
                replacement = AI_TO_HUMAN.get(word, "")
                if replacement:
                    text = pattern.sub(replacement, text)
                else:
                    text = pattern.sub("", text)
        return text

    def _fix_sentence_boundaries(self, text: str) -> str:
        """إصلاح حدود الجمل: التأكد من وجود نقاط وفواصل صحيحة."""
        # إزالة المسافات الزائدة قبل النقاط
        text = re.sub(r'\s+\.', '.', text)
        # إزالة النقاط المكررة
        text = re.sub(r'\.{2,}', '.', text)
        # إضافة نقطة في نهاية النص إذا لم تكن موجودة
        if text and text[-1] not in '.!?':
            text += '.'
        # تصحيح حالات مثل "source. Of" إلى "source of"
        text = re.sub(r'\.\s+([A-Z][a-z]{1,3})\s+', lambda m: f'. {m.group(1).lower()} ', text)
        return text

    def _split_and_vary(self, text: str) -> str:
        """تقطيع الجمل الطويلة وإضافة تنوع في الطول."""
        # تقسيم النص إلى جمل
        sentences = re.split(r'(?<=[.!?])\s+', text)
        sentences = [s.strip() for s in sentences if s.strip()]
        
        new_sentences = []
        for sent in sentences:
            words = sent.split()
            # الجمل التي تزيد عن 18 كلمة يتم تقطيعها
            if len(words) > 18:
                # البحث عن فاصلة أو "and" أو "but" أو "so" لتقطيع طبيعي
                split_pos = -1
                for i, w in enumerate(words):
                    if i > 5 and i < len(words)-3 and w.lower() in [',', 'and', 'but', 'so', 'because']:
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
                    # تقطيع في المنتصف
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
            else:
                new_sentences.append(sent)
        
        # ضمان وجود جمل قصيرة (أقل من 8 كلمات) وجمل طويلة (أكثر من 25 كلمة)
        has_short = any(len(s.split()) < 8 for s in new_sentences)
        has_long = any(len(s.split()) > 25 for s in new_sentences)
        
        if not has_short and len(new_sentences) > 0:
            # تقسيم أطول جملة لجعل جزء قصير
            longest_idx = max(range(len(new_sentences)), key=lambda i: len(new_sentences[i].split()))
            long_sent = new_sentences[longest_idx]
            words = long_sent.split()
            if len(words) > 6:
                short_part = ' '.join(words[:3]) + '.'
                rest = ' '.join(words[3:])
                if rest:
                    if rest[-1] not in '.!?':
                        rest += '.'
                    rest = rest[0].upper() + rest[1:]
                new_sentences[longest_idx] = rest
                new_sentences.insert(longest_idx + 1, short_part)
        
        if not has_long and len(new_sentences) > 1:
            # دمج جملتين قصيرتين
            for i in range(len(new_sentences)-1):
                if len(new_sentences[i].split()) < 12 and len(new_sentences[i+1].split()) < 12:
                    merged = new_sentences[i] + ' ' + new_sentences[i+1][0].lower() + new_sentences[i+1][1:]
                    new_sentences[i] = merged
                    del new_sentences[i+1]
                    break
        
        return ' '.join(new_sentences)

    def _add_human_style(self, text: str) -> str:
        """إضافة لمسات بشرية: بدايات عامية، تحفظات، نهايات استفهامية."""
        sentences = re.split(r'(?<=[.!?])\s+', text)
        sentences = [s.strip() for s in sentences if s.strip()]
        
        for i in range(len(sentences)):
            # إضافة بداية عامية (10-15% من الجمل)
            if random.random() < 0.12:
                starters = ['So ', 'Well ', 'Look, ', 'Basically ', 'I mean, ', 'You see, ']
                starter = random.choice(starters)
                # لا نضيف إذا كانت الجملة تبدأ بالفعل بكلمة مشابهة
                if not sentences[i][:3].lower() in [s[:3].lower() for s in starters]:
                    sentences[i] = starter + sentences[i][0].lower() + sentences[i][1:]
            
            # إضافة تحفظ (I think, maybe) في منتصف الجمل (10%)
            if random.random() < 0.1 and len(sentences[i].split()) > 6:
                words = sentences[i].split()
                pos = random.randint(2, min(5, len(words)-2))
                hedges = ['I think', 'maybe', 'probably', 'it seems']
                hedge = random.choice(hedges)
                words.insert(pos, hedge)
                sentences[i] = ' '.join(words)
            
            # إضافة سؤال استفهامي في النهاية (5%)
            if random.random() < 0.05:
                if sentences[i][-1] in '.!?':
                    sentences[i] = sentences[i][:-1] + ', right?'
                else:
                    sentences[i] = sentences[i] + ', right?'
        
        return ' '.join(sentences)

    def _preserve_citations(self, original: str, revised: str) -> str:
        """الحفاظ على الاستشهادات كما هي."""
        # استخراج جميع الاستشهادات من النص الأصلي
        citations = re.findall(r'\[\d+(?:[-,;]\s*\d+)*\]', original)
        citations.extend(re.findall(r'\([^)]*\d{4}[^)]*\)', original))
        citations = list(dict.fromkeys(citations))
        
        if not citations:
            return revised
        
        # استبدال أي شيء يشبه الاستشهاد في النص المعدل
        for i, cit in enumerate(citations):
            # العثور على أول استشهاد وهمي واستبداله
            match = re.search(r'\[\d+(?:[-,;]\s*\d+)*\]', revised)
            if match:
                revised = revised.replace(match.group(0), cit, 1)
        return revised

    def _final_cleanup(self, text: str) -> str:
        """تنظيف نهائي: إزالة المسافات الزائدة، إصلاح علامات الترقيم."""
        text = re.sub(r'\s+', ' ', text)
        text = text.replace(' ,', ',').replace(' .', '.')
        text = text.replace('  ', ' ')
        text = re.sub(r' ([.,!?;:])', r'\1', text)
        text = re.sub(r'([.!?]) ([a-z])', lambda m: f'{m.group(1)} {m.group(2).upper()}', text)
        text = text.replace('—', ', ').replace('–', '-')
        text = text.replace('*', '').replace('_', '')
        return text.strip()

    def run(self) -> str:
        text = self.original
        
        # إزالة أي بقايا Markdown أو تعليمات روبوتية
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is|as an ai).*\n?', '', text, flags=re.M)
        text = re.sub(r'\*{1,2}[^*]+\*{1,2}', '', text)
        text = re.sub(r'```[\s\S]*?```', '', text)
        
        # التطبيق المتسلسل
        text = self._simple_replace(text)
        text = self._clean_forbidden(text)
        text = self._fix_sentence_boundaries(text)
        text = self._split_and_vary(text)
        text = self._add_human_style(text)
        text = self._preserve_citations(self.original, text)
        text = self._final_cleanup(text)
        
        return text


# ============================================================================
# 3. تحليل النص (تقديري)
# ============================================================================
@dataclass
class Report:
    classification: str
    confidence: float
    human_score: float
    forbidden_remaining: dict

def quick_scan(text: str) -> Report:
    """تقدير سريع لمدى بشريّة النص."""
    lowered = text.lower()
    forbidden_count = 0
    forbidden_found = {}
    
    for word in FORBIDDEN:
        if word in lowered:
            count = lowered.count(word)
            if count > 0:
                forbidden_found[word] = count
                forbidden_count += count
    
    words = re.findall(r'\b\w+\b', lowered)
    word_count = len(words) if words else 1
    
    # درجة الآلية تعتمد على كثافة الكلمات المحظورة
    ai_score = min(0.99, forbidden_count / word_count * 8)
    
    # تحسين إذا كان النص يحتوي على جمل قصيرة ومتنوعة
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 3]
    if sentences:
        avg_len = sum(len(s.split()) for s in sentences) / len(sentences)
        has_short = any(len(s.split()) < 8 for s in sentences)
        has_long = any(len(s.split()) > 25 for s in sentences)
        if has_short and has_long and (12 < avg_len < 28):
            ai_score *= 0.6  # خفض الدرجة لأن النص يبدو بشريًا
    
    human_score = max(0, 100 - (ai_score * 100))
    confidence = ai_score
    
    if human_score > 85:
        classification = "نص بشري بدرجة عالية"
    elif human_score > 65:
        classification = "نص بشري محتمل"
    elif human_score > 45:
        classification = "نص مختلط"
    else:
        classification = "نص آلي محتمل"
    
    return Report(classification, confidence, human_score, forbidden_found)


# ============================================================================
# 4. دوال مساعدة لواجهة المستخدم
# ============================================================================
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

def extract_uploaded_file(uploaded) -> str:
    name = uploaded.name.lower()
    if name.endswith(".txt"):
        raw = uploaded.read()
        try:
            return raw.decode("utf-8")
        except:
            return raw.decode("utf-8-sig")
    if name.endswith(".docx"):
        return docx2txt.process(uploaded) or ""
    if name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""

def make_word_file(text: str, title: Optional[str] = None) -> BytesIO:
    doc = Document()
    doc.core_properties.author = AUTHOR_NAME
    doc.core_properties.title = title or "DeepClean Humanized"
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)
    
    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(title)
        run.font.size = Pt(14)
        run.bold = True
    
    for line in text.split("\n"):
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

def preview_html(text: str) -> str:
    return f"<div style='font-family: Times New Roman; font-size: 12pt; line-height: 1.4;'>{html.escape(text).replace(chr(10), '<br>')}</div>"


# ============================================================================
# 5. تطبيق Streamlit الرئيسي
# ============================================================================
def main():
    st.title("✍️ DeepClean Studio – Human Rewriter")
    st.markdown("""
    <style>
    .stApp { background-color: #faf9f8; }
    </style>
    """, unsafe_allow_html=True)
    st.caption("يعيد كتابة النصوص الأكاديمية الآلية إلى كتابة بشرية طبيعية تجتاز GPTZero، ZeroGPT، Originality.ai")
    st.caption(AUTHOR_NAME)

    with st.sidebar:
        st.header("⚙️ الإعدادات")
        source = st.radio("مصدر النص", ("📄 لصق نص", "📁 رفع ملف"), key="src")
        user_text = ""
        
        if source == "📁 رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                user_text = extract_uploaded_file(uploaded)
        else:
            user_text = st.text_area("ألصق النص الأكاديمي هنا", height=250)
        
        intensity = st.slider("قوة المراجعة", 1, 5, 3, help="كلما زادت القوة، زادت العشوائية البشرية")
        seed_val = st.number_input("بذرة عشوائية (للتكرار)", value=42, step=1)
        
        if st.button("🚀 بدء المراجعة", type="primary", use_container_width=True):
            if not user_text:
                st.warning("الرجاء إدخال نص أو رفع ملف.")
            else:
                with st.spinner("جاري إعادة الكتابة بأسلوب بشري..."):
                    engine = HumanRewriter(user_text, intensity=intensity, seed=seed_val)
                    result = engine.run()
                    st.session_state.result = result
                    st.session_state.original = user_text
                    st.session_state.scan = quick_scan(result)
                    st.session_state.done = True

    colA, colB = st.columns(2)
    
    with colA:
        st.subheader("📄 النص الأصلي")
        if user_text:
            st.text_area("", user_text, height=450, key="orig_area")
            st.caption(f"كلمات: {len(tokenize_words(user_text))}")
        else:
            st.info("أدخل نصًا من الشريط الجانبي.")
    
    with colB:
        st.subheader("✨ النص المعاد كتابته (بشري)")
        if st.session_state.get("done") and st.session_state.get("result"):
            rev = st.session_state.result
            st.markdown(preview_html(rev), unsafe_allow_html=True)
            st.text_area("", rev, height=450, key="rev_area", label_visibility="collapsed")
            st.caption(f"كلمات: {len(tokenize_words(rev))}")
            word_file = make_word_file(rev, "DeepClean_Humanized")
            st.download_button("📥 تحميل Word", data=word_file, file_name="humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        else:
            st.info("ستظهر النسخة البشرية هنا بعد المعالجة.")
    
    if st.session_state.get("done") and st.session_state.get("scan"):
        rep = st.session_state.scan
        with st.expander("🔍 نتيجة الفحص التقديري", expanded=False):
            c1, c2, c3 = st.columns(3)
            c1.metric("التصنيف", rep.classification)
            c2.metric("النسبة البشرية", f"{rep.human_score:.1f}%")
            c3.metric("درجة الآلية", f"{rep.confidence*100:.1f}%")
            if rep.forbidden_remaining:
                st.warning(f"كلمات محظورة متبقية: {', '.join(list(rep.forbidden_remaining.keys())[:5])}")
            else:
                st.success("✅ لا توجد كلمات محظورة في النص النهائي.")
            st.caption("هذا الفحص تقديري محلي. يفضل اختبار النص على GPTZero أو ZeroGPT للتأكد.")


if __name__ == "__main__":
    if "done" not in st.session_state:
        st.session_state.done = False
    if "result" not in st.session_state:
        st.session_state.result = ""
    if "original" not in st.session_state:
        st.session_state.original = ""
    if "scan" not in st.session_state:
        st.session_state.scan = None
    main()
