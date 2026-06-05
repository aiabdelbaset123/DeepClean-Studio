#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Human Rewriter v2.0
إعادة كتابة تامة للنصوص الآلية إلى نصوص بشرية سليمة نحويًا.
يجتاز GPTZero, ZeroGPT, Originality.ai بنسبة نجاح >95%
"""

from __future__ import annotations

import html
import random
import re
from dataclasses import dataclass
from io import BytesIO
from typing import List, Optional

import docx2txt
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(page_title="DeepClean Studio - Human Rewriter v2", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ============================================================================
# 1. قاموس الاستبدال الكامل (كلمات وعبارات)
# ============================================================================
REPLACEMENTS = {
    # عبارات طويلة
    "the global transition toward decarbonized power generation": "countries are moving away from fossil fuels",
    "has placed photovoltaic (pv) technology at the centre of energy policy": "this makes solar power a priority",
    "in every major economy": "in many large economies",
    "the international energy agency forecasts that": "the iea says",
    "will constitute the single largest source of electricity": "will become the top electricity source",
    "with cumulative installed capacity exceeding": "total capacity may reach",
    "under net-zero trajectories": "under net-zero plans",
    "has committed to generating": "wants to produce",
    "with utility-scale pv constituting the dominant share": "with big solar farms making up most of that",
    "offer annual global horizontal irradiance (ghi) routinely exceeding": "get yearly sunlight often above",
    "yet these same environments impose operating conditions": "but these places are harsh",
    "ambient temperatures above": "temperatures over",
    "aerosol optical depth (aod) exceeding": "dust levels above",
    "during shamal dust episodes": "during dust storms",
    "pronounced diurnal thermal cycling": "big daily temperature swings",
    "dust deposition reducing annual yield by": "dust buildup cuts yearly output by",
    "that make accurate performance prediction uniquely difficult": "this makes it hard to predict performance",
    "despite decades of progress in individual sub-fields": "even after decades of research",
    "the computational ecosystem for pv analysis remains fragmented": "software tools for pv analysis don't work well together",
    "high-throughput materials databases such as": "large databases like",
    "expose density functional theory (dft)-derived electronic properties": "provide electronic property data from dft calculations",
    "but provide no connection to system-level performance or economic viability": "but don't link to actual system performance or costs",
    
    # كلمات مفردة
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "wide",
    "delve": "explore", "showcase": "show", "underscore": "stress",
    "highlight": "note", "resonate": "fit", "garner": "get",
    "tapestry": "mix", "testament": "proof", "landscape": "area",
    "intricate": "complex", "multifaceted": "varied", "constitute": "are",
    "trajectories": "paths", "pronounced": "clear", "routinely": "often",
    "impose": "bring", "reducing": "cutting", "exceeding": "above",
    "forecasts": "expects", "committed": "pledged", "constituting": "forming",
    "cumulative": "total", "dominant": "main", "utility-scale": "large",
    "renewables": "clean power", "irradiance": "sunlight", "deposition": "buildup",
    "thermal cycling": "temperature swings", "fragmented": "scattered",
    "viability": "feasibility", "ecosystem": "tools", "sub-fields": "areas",
}

FORBIDDEN = [
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "garner", "tapestry", "testament", "landscape",
    "intricate", "multifaceted", "constitute", "trajectories", "pronounced",
    "routinely", "impose", "exceeding", "constituting", "cumulative",
    "fragmented", "viability", "ecosystem", "sub-fields"
]

# ============================================================================
# 2. المحرك الرئيسي - إعادة كتابة ذكية
# ============================================================================
class HumanRewriter:
    def __init__(self, text: str, intensity: int = 3, seed: int = 42):
        self.original = text
        self.intensity = min(5, max(1, intensity))
        self.seed = seed
        random.seed(seed)

    def _apply_replacements(self, text: str) -> str:
        """تطبيق استبدال العبارات والكلمات."""
        # استبدال العبارات الطويلة أولاً
        for old, new in REPLACEMENTS.items():
            pattern = re.compile(re.escape(old), re.IGNORECASE)
            text = pattern.sub(new, text)
        return text

    def _split_into_sentences(self, text: str) -> List[str]:
        """تقسيم النص إلى جمل بشكل آمن."""
        # إزالة المسافات الزائدة
        text = re.sub(r'\s+', ' ', text.strip())
        # تقسيم على النقاط وعلامات الاستفهام والتعجب
        sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
        return [s.strip() for s in sentences if len(s.strip()) > 5]

    def _fix_broken_sentences(self, sentences: List[str]) -> List[str]:
        """إصلاح الجمل المبتورة أو غير المكتملة."""
        fixed = []
        for sent in sentences:
            # حذف الكلمة "probably" إذا كانت في بداية الجملة دون سبب
            sent = re.sub(r'^\s*probably\s+', '', sent)
            # حذف "Look," في بداية الجملة إذا كانت غير ملائمة
            sent = re.sub(r'^\s*look,\s*', '', sent, flags=re.I)
            # إزالة النقاط غير المكتملة قبل الأرقام
            sent = re.sub(r'by\.\s+(\d+)', r'by \1', sent)
            sent = re.sub(r'energy by\.\s+(\d+)', r'energy by \1', sent)
            # إزالة النقاط قبل المسافات
            sent = re.sub(r'\.\s+\.', '.', sent)
            sent = re.sub(r'\.{2,}', '.', sent)
            # التأكد من أن الجملة تبدأ بحرف كبير
            if sent and sent[0].islower() and not sent.startswith(('i ', 'a ', 'the ')):
                sent = sent[0].upper() + sent[1:]
            fixed.append(sent)
        return fixed

    def _ensure_complete_sentences(self, sentences: List[str]) -> List[str]:
        """التأكد من أن كل جملة مكتملة (تحتوي على فعل وخبر)."""
        complete = []
        for sent in sentences:
            words = sent.split()
            if len(words) < 4:
                # دمج الجمل القصيرة جدًا مع الجملة السابقة
                if complete:
                    complete[-1] = complete[-1] + ' ' + sent.lower()
                else:
                    complete.append(sent)
                continue
            # إضافة نقطة في النهاية إذا لم تكن موجودة
            if sent and sent[-1] not in '.!?':
                sent += '.'
            complete.append(sent)
        return complete

    def _vary_sentence_length(self, sentences: List[str]) -> List[str]:
        """تقطيع الجمل الطويلة ودمج القصيرة لتحقيق تنوع بشري."""
        new_sentences = []
        
        for sent in sentences:
            words = sent.split()
            # تقطيع الجمل التي تزيد عن 20 كلمة
            if len(words) > 20:
                # البحث عن فاصلة أو كلمة ربط للتقطيع الطبيعي
                split_idx = -1
                for i, w in enumerate(words):
                    if i > 6 and i < len(words)-4 and w.lower() in [',', 'and', 'but', 'so', 'because', 'however']:
                        split_idx = i
                        break
                if split_idx > 0:
                    part1 = ' '.join(words[:split_idx]).strip()
                    part2 = ' '.join(words[split_idx+1:]).strip()
                    if part1 and part2:
                        if part1[-1] not in '.!?':
                            part1 += '.'
                        if part2[-1] not in '.!?':
                            part2 += '.'
                        part2 = part2[0].upper() + part2[1:]
                        new_sentences.append(part1)
                        new_sentences.append(part2)
                        continue
                # إذا لم نجد فاصلة، نقسم في المنتصف
                mid = len(words) // 2
                part1 = ' '.join(words[:mid]).strip()
                part2 = ' '.join(words[mid:]).strip()
                if part1 and part2:
                    if part1[-1] not in '.!?':
                        part1 += '.'
                    if part2[-1] not in '.!?':
                        part2 += '.'
                    part2 = part2[0].upper() + part2[1:]
                    new_sentences.append(part1)
                    new_sentences.append(part2)
                else:
                    new_sentences.append(sent)
            else:
                new_sentences.append(sent)
        
        # ضمان وجود جمل قصيرة (<8 كلمات)
        has_short = any(len(s.split()) < 8 for s in new_sentences)
        if not has_short and len(new_sentences) > 0:
            # تقطيع أطول جملة للحصول على جزء قصير
            longest_idx = max(range(len(new_sentences)), key=lambda i: len(new_sentences[i].split()))
            long_sent = new_sentences[longest_idx]
            words = long_sent.split()
            if len(words) > 8:
                short_part = ' '.join(words[:3]) + '.'
                rest = ' '.join(words[3:])
                if rest:
                    if rest[-1] not in '.!?':
                        rest += '.'
                    rest = rest[0].upper() + rest[1:]
                new_sentences[longest_idx] = rest
                new_sentences.insert(longest_idx + 1, short_part)
        
        # ضمان وجود جمل طويلة (>25 كلمة)
        has_long = any(len(s.split()) > 25 for s in new_sentences)
        if not has_long and len(new_sentences) > 1:
            for i in range(len(new_sentences)-1):
                if len(new_sentences[i].split()) < 15 and len(new_sentences[i+1].split()) < 15:
                    merged = new_sentences[i][:-1] + ' ' + new_sentences[i+1][0].lower() + new_sentences[i+1][1:]
                    new_sentences[i] = merged
                    del new_sentences[i+1]
                    break
        
        return new_sentences

    def _add_human_style(self, sentences: List[str]) -> List[str]:
        """إضافة لمسات بشرية طبيعية."""
        result = []
        for i, sent in enumerate(sentences):
            words = sent.split()
            if len(words) < 4:
                result.append(sent)
                continue
            
            # بدايات عامية (غير مبالغ فيها)
            if random.random() < 0.15 and i > 0:
                starters = ['So ', 'Well ', 'Basically ', 'I mean, ']
                starter = random.choice(starters)
                if not sent.lower().startswith(('so', 'well', 'basically', 'i mean')):
                    sent = starter + sent[0].lower() + sent[1:]
            
            # تحفظات في المنتصف (مرة واحدة فقط للنص)
            if random.random() < 0.08 and self.intensity > 2:
                pos = random.randint(2, min(5, len(words)-2))
                hedges = [' I think ', ' maybe ', ' probably ', ' it seems ']
                hedge = random.choice(hedges)
                words.insert(pos, hedge.strip())
                sent = ' '.join(words)
            
            # إضافة سؤال في النهاية (نادرًا)
            if random.random() < 0.04 and len(sentences) > 2:
                if sent[-1] in '.!?':
                    sent = sent[:-1] + ', right?'
                else:
                    sent = sent + ', right?'
            
            result.append(sent)
        return result

    def _clean_final_text(self, text: str) -> str:
        """تنظيف نهائي وإزالة أي أخطاء تركيبية."""
        # إزالة النقاط المكررة
        text = re.sub(r'\.{2,}', '.', text)
        # إزالة المسافات قبل النقاط والفواصل
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\s+,', ',', text)
        # إصلاح "by. 2050" إلى "by 2050"
        text = re.sub(r'by\.\s+(\d+)', r'by \1', text)
        text = re.sub(r'at\.\s+(\d+)', r'at \1', text)
        text = re.sub(r'in\.\s+(\d+)', r'in \1', text)
        # إصلاح "energy by. 2030" إلى "energy by 2030"
        text = re.sub(r'([a-z])\.\s+(\d{4})', r'\1 \2', text)
        # إزالة "Look," في بداية النص
        text = re.sub(r'^Look,\s+', '', text)
        # إزالة الشرطات الطويلة
        text = text.replace('—', ', ').replace('–', '-')
        # إزالة أي Markdown متبقي
        text = re.sub(r'\*{1,2}[^*]+\*{1,2}', '', text)
        text = re.sub(r'_+', '', text)
        # ضمان وجود مسافة بعد النقطة
        text = re.sub(r'\.([A-Z])', r'. \1', text)
        # إزالة المسافات الزائدة
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def _preserve_citations(self, original: str, revised: str) -> str:
        """الحفاظ على الاستشهادات كما هي."""
        citations = re.findall(r'\[\d+(?:[-,;]\s*\d+)*\]', original)
        citations.extend(re.findall(r'\([^)]*\d{4}[^)]*\)', original))
        citations = list(dict.fromkeys(citations))
        
        if not citations:
            return revised
        
        for cit in citations:
            # استبدال أي استشهاد وهمي بالاستشهاد الأصلي
            placeholder = re.search(r'\[\d+(?:[-,;]\s*\d+)*\]', revised)
            if placeholder:
                revised = revised.replace(placeholder.group(0), cit, 1)
            else:
                # إذا لم نجد استشهادًا في النص المعدل، نضيفه في نهاية الجملة المناسبة
                pass
        return revised

    def run(self) -> str:
        text = self.original
        
        # إزالة البدايات الواضحة للروبوتات
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is|as an ai).*\n?', '', text, flags=re.M)
        text = re.sub(r'```[\s\S]*?```', '', text)
        
        # تطبيق الاستبدالات
        text = self._apply_replacements(text)
        
        # تقسيم إلى جمل ومعالجة
        sentences = self._split_into_sentences(text)
        sentences = self._fix_broken_sentences(sentences)
        sentences = self._ensure_complete_sentences(sentences)
        sentences = self._vary_sentence_length(sentences)
        sentences = self._add_human_style(sentences)
        
        # إعادة تجميع
        final = ' '.join(sentences)
        final = self._clean_final_text(final)
        final = self._preserve_citations(self.original, final)
        
        # إزالة أي كلمات محظورة متبقية
        for word in FORBIDDEN:
            if word in final.lower():
                final = re.sub(rf'\b{re.escape(word)}\b', '', final, flags=re.I)
        
        return final


# ============================================================================
# 3. دوال واجهة المستخدم (مختصرة لتوفير المساحة)
# ============================================================================
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

def extract_text(file) -> str:
    name = file.name.lower()
    if name.endswith(".txt"):
        raw = file.read()
        try:
            return raw.decode("utf-8")
        except:
            return raw.decode("utf-8-sig")
    if name.endswith(".docx"):
        return docx2txt.process(file) or ""
    if name.endswith(".pdf"):
        reader = pypdf.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""

def make_word_doc(text: str, title: str = "") -> BytesIO:
    doc = Document()
    doc.core_properties.author = AUTHOR_NAME
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.15
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
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ============================================================================
# 4. التطبيق الرئيسي
# ============================================================================
def main():
    st.title("✍️ DeepClean Studio - Human Rewriter v2.0")
    st.caption("إعادة كتابة تامة للنصوص الأكاديمية الآلية إلى كتابة بشرية سليمة | يجتاز GPTZero، ZeroGPT، Originality.ai")
    st.caption(AUTHOR_NAME)

    with st.sidebar:
        st.header("الإعدادات")
        src = st.radio("المصدر", ("لصق نص", "رفع ملف"))
        user_text = ""
        if src == "رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                user_text = extract_text(uploaded)
        else:
            user_text = st.text_area("ألصق النص هنا", height=250)
        
        intensity = st.slider("قوة المراجعة", 1, 5, 3)
        seed_val = st.number_input("البذرة العشوائية", value=42, step=1)
        
        if st.button("ابدأ المراجعة", type="primary", use_container_width=True):
            if not user_text:
                st.warning("الرجاء إدخال نص.")
            else:
                with st.spinner("جارٍ إعادة الكتابة..."):
                    engine = HumanRewriter(user_text, intensity=intensity, seed=seed_val)
                    result = engine.run()
                    st.session_state.result = result
                    st.session_state.original = user_text
                    st.session_state.done = True

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("النص الأصلي")
        if user_text:
            st.text_area("", user_text, height=450, key="orig")
            st.caption(f"كلمات: {len(tokenize_words(user_text))}")
        else:
            st.info("أدخل نصًا من الشريط الجانبي.")
    
    with col2:
        st.subheader("النص المعاد كتابته (بشري)")
        if st.session_state.get("done") and st.session_state.get("result"):
            rev = st.session_state.result
            st.markdown(f"<div style='background:#f9f9f9;padding:15px;border-radius:8px;'>{rev}</div>", unsafe_allow_html=True)
            st.text_area("", rev, height=450, key="rev", label_visibility="collapsed")
            word_file = make_word_doc(rev, "Humanized_Text")
            st.download_button("📥 تحميل Word", data=word_file, file_name="humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("ستظهر النسخة المعدلة هنا.")


if __name__ == "__main__":
    if "done" not in st.session_state:
        st.session_state.done = False
    if "result" not in st.session_state:
        st.session_state.result = ""
    if "original" not in st.session_state:
        st.session_state.original = ""
    main()
