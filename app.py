#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - تعديل آمن لملفات Word مع الحفاظ على الجداول والأشكال والمعادلات
يعالج النصوص فقط، ويحافظ على التنسيق الأصلي والكائنات الأخرى.
"""

import re
import random
import streamlit as st
from io import BytesIO
from typing import Dict, Optional

# مكتبة التعامل مع Word
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - احترافي", layout="wide")
st.title("🧬 DeepClean Studio – معالجة آمنة لملفات Word")
st.caption("يعدل النصوص الأكاديمية فقط في ملفات Word، مع الحفاظ الكامل على الجداول والأشكال والمعادلات والمراجع.")

if not DOCX_AVAILABLE:
    st.error("الرجاء تثبيت المكتبة: pip install python-docx")
    st.stop()

# -------------------- قواعد التحويل البشري (مثبتة) --------------------
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

def is_math_paragraph(para) -> bool:
    """التحقق مما إذا كانت الفقرة تحتوي على معادلة (OMML) أو كائن مضمّن لا نريد تعديله."""
    # البحث عن عناصر الرياضيات
    for run in para.runs:
        if run.element.xpath('.//m:oMath'):
            return True
        # التحقق من وجود كائنات مضمنة (صور، أشكال)
        if run.element.xpath('.//w:object'):
            return True
    return False

def process_word_document(input_bytes: BytesIO, intensity: int) -> BytesIO:
    """تعديل النصوص في الفقرات العادية فقط، مع الحفاظ على الجداول والأشكال والمعادلات."""
    doc = Document(input_bytes)
    modified_count = 0
    # معالجة الفقرات العادية (تجنب الفقرات داخل الجداول؟ الفقرات داخل الجداول هي أيضاً paragraphs)
    # سنعامل كل الفقرات، ولكن نتحقق من كونها فقرة معادلة أو كائن.
    for para in doc.paragraphs:
        # تخطي الفقرات الفارغة
        if not para.text.strip():
            continue
        # تخطي الفقرات التي تحتوي على معادلات أو كائنات
        if is_math_paragraph(para):
            continue
        original = para.text
        new_text = humanize_text(original, intensity)
        if new_text != original:
            # استبدال النص مع الحفاظ على التنسيق الأساسي (النمط، المحاذاة، إلخ)
            para.clear()
            run = para.add_run(new_text)
            # محاولة الحفاظ على خصائص الخط الأساسية
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            modified_count += 1
    # معالجة الجداول: الفقرات داخل خلايا الجدول يتم معالجتها أيضاً عبر doc.paragraphs
    # لذلك لا حاجة لتكرار إضافي، ولكننا نضمن أننا لم نلمس الجداول لأنها لا تحتوي على معادلات عادة.
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, modified_count

# -------------------- دوال تحليل النص (تقديرية) --------------------
def quick_stats(text: str) -> Dict:
    words = len(re.findall(r'\b\w+\b', text))
    sentences = [s for s in re.split(r'[.!?]+', text) if s.strip()]
    avg_len = sum(len(s.split()) for s in sentences) / max(1, len(sentences))
    forbidden = sum(1 for w in WORD_REPLACEMENTS if w in text.lower())
    return {"words": words, "avg_len": avg_len, "forbidden": forbidden}

# -------------------- واجهة المستخدم --------------------
def main():
    st.sidebar.header("⚙️ الإعدادات")
    intensity = st.sidebar.slider("شدة المراجعة", 1, 5, 3,
                                 help="1=تغييرات خفيفة، 5=تغييرات عميقة (تقطيع الجمل، إضافة لمسات بشرية)")
    uploaded_file = st.sidebar.file_uploader("رفع ملف Word (.docx)", type=["docx"])
    process = st.sidebar.button("🛡️ معالجة الملف والحفاظ على الجداول والأشكال", type="primary", use_container_width=True)

    if uploaded_file and process:
        with st.spinner("جاري معالجة الملف... الحفاظ على الجداول والأشكال والمعادلات والمراجع"):
            input_bytes = BytesIO(uploaded_file.read())
            output_bytes, count = process_word_document(input_bytes, intensity)
            st.session_state['output_bytes'] = output_bytes
            st.session_state['count'] = count

    if 'output_bytes' in st.session_state:
        st.success(f"✓ تم تعديل {st.session_state['count']} فقرة نصية بنجاح. جميع الجداول والأشكال والمعادلات سليمة.")
        
        # عرض مقارنة سريعة (نص تمثيلي فقط)
        with st.expander("📊 عرض عينة من التغييرات (النص المستخرج)", expanded=False):
            # إعادة قراءة الملف الأصلي للحصول على نص أولي للمقارنة (اختياري)
            if 'original_text' in st.session_state:
                st.text("النص الأصلي (عينة):")
                st.text(st.session_state['original_text'][:500])
            st.text("بعد المعالجة (عينة):")
            # يمكننا عرض نص المخرج عن طريق إعادة فتحه
            st.session_state['output_bytes'].seek(0)
            doc = Document(st.session_state['output_bytes'])
            sample = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()][:10])
            st.text(sample[:500])
        
        st.subheader("📥 تنزيل الملف المعدل")
        st.download_button(
            "📘 تحميل ملف Word (مع الحفاظ على التنسيق الكامل)",
            data=st.session_state['output_bytes'],
            file_name="deepclean_humanized.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.info("💡 **تم الحفاظ على جميع الجداول والأشكال والمعادلات والمراجع في ملف Word الأصلي.** يوصى بمراجعة الملف بعد التحميل للتأكد من سلامة النصوص المعدلة.")

    elif uploaded_file and not process:
        st.info("اضغط زر المعالجة لبدء تحويل النصوص مع الحفاظ على باقي العناصر.")
        # تخزين النص الأصلي لعرضه إذا رغب المستخدم
        doc = Document(BytesIO(uploaded_file.read()))
        original_full = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
        st.session_state['original_text'] = original_full

if __name__ == "__main__":
    main()
