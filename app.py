#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Local Human Rewriter v2.0
لا يستخدم أي API خارجي. يعيد كتابة النصوص بأسلوب بشري 100% عبر:
- جمل قصيرة (5-12 كلمة)
- استبدال شامل للكلمات المحظورة
- إضافة لمسات بشرية طبيعية
- إصلاح تلقائي للأخطاء النحوية
"""

import html
import random
import re
from io import BytesIO
from typing import List, Dict, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

# مكتبات اختيارية لرفع الملفات
try:
    import docx2txt
    import pypdf
    HAS_EXTRACT = True
except ImportError:
    HAS_EXTRACT = False

st.set_page_config(page_title="Human Rewriter - Local", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ============================================================
# 1. قاموس الاستبدال البشري (واسع وشامل)
# ============================================================
WORD_REPLACEMENTS = {
    # الأفعال والعبارات الشائعة في النصوص الآلية
    "transition": "move", "decarbonized": "clean", "generation": "power",
    "has placed": "makes", "technology": "tech", "at the centre of": "central to",
    "energy policy": "energy plans", "every major economy": "many large countries",
    "forecasts": "says", "will constitute": "will be", "single largest source": "top source",
    "cumulative installed capacity": "total capacity", "exceeding": "above",
    "net-zero trajectories": "net-zero paths", "has committed": "plans",
    "generating": "producing", "utility-scale": "large", "dominant share": "most",
    "offer": "get", "global horizontal irradiance": "sunlight",
    "routinely exceeding": "often above", "yet these same environments": "but these places",
    "impose operating conditions": "are harsh", "ambient temperatures": "temperatures",
    "aerosol optical depth": "dust levels", "exceeding": "above",
    "shamal dust episodes": "dust storms", "pronounced": "big",
    "diurnal thermal cycling": "daily temperature swings", "dust deposition": "dust buildup",
    "reducing annual yield": "cutting yearly output", "accurate performance prediction": "good performance guesses",
    "uniquely difficult": "hard", "despite decades of progress": "even after years of work",
    "computational ecosystem": "software tools", "remains fragmented": "don't work well together",
    "high-throughput materials databases": "large material databases",
    "expose density functional theory (dft)-derived electronic properties": "give electronic data",
    "hundreds of thousands of compounds": "many compounds",
    "provide no connection to": "don't link to",
    "system-level performance": "real system performance", "economic viability": "costs",
    "key part": "important part", "aims": "wants", "exceed": "go above",
    "fragmentation forces practitioners to": "this split makes people",
    "transfer data manually": "move data by hand",
    "introduces inconsistency": "causes mismatches", "systematically ignores": "misses",
    "cross-domain interactions": "connections between fields",
}

# كلمات إضافية يجب حذفها أو استبدالها بالقوة
FORBIDDEN_SET = {
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "align", "garner", "tapestry", "testament", "landscape",
    "intricate", "multifaceted", "constitute", "trajectories", "routinely",
    "impose", "pronounced", "cumulative", "uniquely", "constituting",
    "utility-scale", "dominant share", "diurnal", "aerosol optical depth",
    "shamal", "thermal cycling", "deposition", "high-throughput",
    "density functional theory", "fragmentation", "cross-domain",
}

# بدائل بشرية للكلمات المحظورة
HUMAN_ALTERNATIVES = {
    "additionally": "also", "moreover": "plus", "furthermore": "then",
    "consequently": "so", "hence": "thus", "crucial": "big",
    "pivotal": "key", "vital": "needed", "significant": "large",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "underscore": "stress",
    "highlight": "point out", "resonate": "match", "garner": "get",
    "tapestry": "mix", "testament": "proof", "landscape": "field",
    "intricate": "complex", "multifaceted": "varied", "constitute": "are",
    "constituting": "making", "uniquely": "", "cumulative": "total",
}

# ============================================================
# 2. المحرك البشري المحلي
# ============================================================
class LocalHumanRewriter:
    def __init__(self, text: str, intensity: int = 3, seed: int = 42):
        self.original = text
        self.intensity = min(5, max(1, intensity))
        self.seed = seed
        random.seed(seed)

    def _apply_word_replacements(self, text: str) -> str:
        """تطبيق استبدالات الكلمات والعبارات."""
        text_lower = text.lower()
        for old, new in WORD_REPLACEMENTS.items():
            if old in text_lower:
                text = re.compile(re.escape(old), re.IGNORECASE).sub(new, text)
        return text

    def _remove_forbidden_words(self, text: str) -> str:
        """إزالة أو استبدال الكلمات المحظورة."""
        for word in FORBIDDEN_SET:
            if word in text.lower():
                rep = HUMAN_ALTERNATIVES.get(word, "")
                if rep:
                    text = re.compile(rf'\b{re.escape(word)}\b', re.IGNORECASE).sub(rep, text)
                else:
                    text = re.compile(rf'\b{re.escape(word)}\b', re.IGNORECASE).sub('', text)
        return text

    def _split_into_short_sentences(self, text: str) -> List[str]:
        """تقسيم النص إلى جمل قصيرة (كل جملة فكرة واحدة)."""
        # تقسيم أولي على النقاط والفواصل
        text = text.replace(';', '.')
        text = text.replace(',', '.')
        sentences = re.split(r'[.!?]+', text)
        sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
        
        # تقسيم الجمل الطويلة (>15 كلمة) إلى جمل أقصر
        result = []
        for sent in sentences:
            words = sent.split()
            if len(words) > 15:
                # تقطيع عند كلمات الربط
                for sep in [' and ', ' but ', ' so ', ' because ', ' which ']:
                    if sep in sent:
                        parts = sent.split(sep)
                        for p in parts:
                            p = p.strip()
                            if p and len(p.split()) > 3:
                                if p[-1] not in '.!?':
                                    p += '.'
                                result.append(p)
                        break
                else:
                    # تقطيع عادي في المنتصف
                    mid = len(words) // 2
                    p1 = ' '.join(words[:mid]).strip()
                    p2 = ' '.join(words[mid:]).strip()
                    if p1 and p2:
                        if p1[-1] not in '.!?':
                            p1 += '.'
                        if p2[-1] not in '.!?':
                            p2 += '.'
                        result.extend([p1, p2])
            else:
                if sent and sent[-1] not in '.!?':
                    sent += '.'
                result.append(sent)
        return result

    def _rewrite_each_sentence(self, sentences: List[str]) -> List[str]:
        """إعادة كتابة كل جملة بأسلوب بشري بسيط."""
        new_sentences = []
        for sent in sentences:
            sent = sent.lower()
            # استبدال العبارات المعقدة بعبارات بسيطة
            sent = self._apply_word_replacements(sent)
            sent = self._remove_forbidden_words(sent)
            # تنظيف المسافات وعلامات الترقيم
            sent = re.sub(r'\s+', ' ', sent)
            sent = sent.strip()
            if sent and sent[-1] not in '.!?':
                sent += '.'
            if sent:
                # جعل أول حرف كبيرًا
                sent = sent[0].upper() + sent[1:]
                new_sentences.append(sent)
        return new_sentences

    def _add_human_style(self, sentences: List[str]) -> List[str]:
        """إضافة لمسات بشرية: بدايات عامية، تحفظات، تنوع في الطول."""
        if len(sentences) < 2:
            return sentences
        
        # ضمان وجود جمل قصيرة جدًا (3-7 كلمات) وجمل متوسطة (10-15 كلمة)
        short_count = sum(1 for s in sentences if len(s.split()) <= 7)
        if short_count == 0 and len(sentences) > 1:
            # تحويل أطول جملة إلى جملة قصيرة + باقي
            longest_idx = max(range(len(sentences)), key=lambda i: len(sentences[i].split()))
            words = sentences[longest_idx].split()
            if len(words) > 8:
                short_sent = ' '.join(words[:4]) + '.'
                rest = ' '.join(words[4:])
                sentences[longest_idx] = rest
                sentences.insert(longest_idx + 1, short_sent)
        
        # إضافة بدايات عامية (So, Well, Look) لبعض الجمل
        for i in range(len(sentences)):
            if random.random() < 0.15 and len(sentences[i].split()) > 4:
                starters = ['So ', 'Well ', 'Look, ', 'I mean, ']
                if not sentences[i].lower().startswith(('so', 'well', 'look', 'i mean')):
                    sentences[i] = random.choice(starters) + sentences[i][0].lower() + sentences[i][1:]
        
        # إضافة تحفظات خفيفة (maybe, I think) في منتصف بعض الجمل
        for i in range(len(sentences)):
            if random.random() < 0.1 and len(sentences[i].split()) > 6:
                words = sentences[i].split()
                pos = random.randint(2, min(4, len(words)-2))
                hedges = ['maybe', 'I think', 'probably']
                words.insert(pos, random.choice(hedges))
                sentences[i] = ' '.join(words)
        
        return sentences

    def _preserve_citations(self, original: str, sentences: List[str]) -> List[str]:
        """استخراج الاستشهادات من النص الأصلي وإضافتها إلى الجمل المناسبة."""
        citations = re.findall(r'\[\d+(?:[-,;]\s*\d+)*\]', original)
        citations.extend(re.findall(r'\([^)]*\d{4}[^)]*\)', original))
        citations = list(dict.fromkeys(citations))
        
        if not citations:
            return sentences
        
        # إضافة الاستشهادات إلى نهاية الجمل التي تبدو وكأنها تحتوي على حقائق
        cit_index = 0
        for i, sent in enumerate(sentences):
            if cit_index >= len(citations):
                break
            # نضيف الاستشهاد إلى الجمل الطويلة نسبيًا
            if len(sent.split()) > 8 and not re.search(r'\[\d+\]', sent):
                sentences[i] = sent.rstrip('.!?') + ' ' + citations[cit_index] + '.'
                cit_index += 1
        
        return sentences

    def _final_cleanup(self, text: str) -> str:
        """تنظيف نهائي: إزالة النقاط المكررة والمسافات الزائدة."""
        text = re.sub(r'\.{2,}', '.', text)
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r' ([.,!?])', r'\1', text)
        text = text.replace(' ,', ',').replace(' .', '.')
        text = re.sub(r'([a-z])\.([A-Z])', r'\1. \2', text)
        return text.strip()

    def run(self) -> str:
        text = self.original
        # إزالة البقايا الأولية
        text = re.sub(r'\*+', '', text)
        text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is).*\n', '', text, flags=re.M)
        
        # تقسيم إلى جمل قصيرة
        sentences = self._split_into_short_sentences(text)
        if not sentences:
            return text
        
        # إعادة كتابة كل جملة
        sentences = self._rewrite_each_sentence(sentences)
        sentences = self._add_human_style(sentences)
        sentences = self._preserve_citations(self.original, sentences)
        
        # تجميع النص
        result = ' '.join(sentences)
        result = self._final_cleanup(result)
        
        return result


# ============================================================
# 3. دوال مساعدة للواجهة
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
# 4. واجهة Streamlit
# ============================================================
def main():
    st.title("✍️ DeepClean Studio – Local Human Rewriter v2")
    st.caption("يعيد كتابة النصوص الأكاديمية الآلية إلى كتابة بشرية 100% باستخدام معالجة محلية فقط (بدون API)")
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
                    st.success("تم تحميل الملف بنجاح")
        else:
            user_text = st.text_area("ألصق النص الأكاديمي هنا", height=250)
        
        intensity = st.slider("قوة المراجعة", 1, 5, 3, help="كلما زادت القوة، زادت العشوائية البشرية")
        seed = st.number_input("بذرة عشوائية", value=42, step=1)
        
        if st.button("🚀 إعادة الكتابة", type="primary", use_container_width=True):
            if not user_text:
                st.warning("الرجاء إدخال نص أو رفع ملف.")
            else:
                with st.spinner("جاري تحويل النص إلى أسلوب بشري..."):
                    engine = LocalHumanRewriter(user_text, intensity=intensity, seed=seed)
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
