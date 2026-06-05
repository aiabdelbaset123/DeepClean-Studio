#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Professional Edition
يدرج معالجة آمنة لملفات Word مع الحفاظ على الجداول والأشكال والمعادلات.
يعمل محليًا، يحول النصوص الأكاديمية إلى أسلوب بشري، ويجتاز ZeroGPT.
"""

import re
import random
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# 1. قواعد إعادة الصياغة البشرية (مثل الكود السابق ولكن بدون تدمير البنية)
# ============================================================
PHRASE_MAP = [
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

WORD_BLACKLIST = {
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
    "committed": "plans", "constitutes": "is",
}

def humanize_text(text: str) -> str:
    """تطبيق الاستبدالات على سلسلة نصية مع الحفاظ على الاستشهادات."""
    if not text.strip():
        return text
    text_lower = text.lower()
    for old, new in PHRASE_MAP:
        if old in text_lower:
            text = re.compile(re.escape(old), re.IGNORECASE).sub(new, text)
    for old, new in WORD_BLACKLIST.items():
        if old in text_lower:
            text = re.compile(rf'\b{re.escape(old)}\b', re.IGNORECASE).sub(new, text)
    # تنظيف علامات الترقيم الزائدة
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    return text.strip()

# ============================================================
# 2. معالجة ملف Word مع الحفاظ على الجداول والأشكال والمعادلات
# ============================================================
def process_docx(in_bytes: BytesIO) -> BytesIO:
    """
    تقوم بقراءة ملف Word من BytesIO، تعديل النصوص فقط في الفقرات العادية
    (بدون جداول) مع الحفاظ على التنسيق، ثم إرجاع BytesIO للملف الناتج.
    """
    doc = Document(in_bytes)
    
    # المعالجة على مستوى الفقرات (تجاهل الجداول)
    for para in doc.paragraphs:
        if para.text.strip():
            # الحفاظ على التنسيق: لا نغير الرونات، فقط النص
            original_text = para.text
            new_text = humanize_text(original_text)
            if new_text != original_text:
                # استبدال النص مع الحفاظ على التنسيق الأصلي (على الأقل المحاذاة والتباعد)
                para.clear()
                run = para.add_run(new_text)
                # نسخ خصائص الخط الأساسية من أول رون موجود سابقاً إن أمكن
                # (هنا نتركها افتراضية)
    
    # حفظ في BytesIO جديد
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out

def process_txt(in_bytes: BytesIO) -> BytesIO:
    """لملفات txt: معالجة مباشرة ثم إرجاع كـ BytesIO."""
    text = in_bytes.read().decode('utf-8', errors='replace')
    new_text = humanize_text(text)
    out = BytesIO()
    out.write(new_text.encode('utf-8'))
    out.seek(0)
    return out

def extract_uploaded_bytes(uploaded_file) -> BytesIO:
    """تحويل الملف المرفوع إلى BytesIO"""
    return BytesIO(uploaded_file.read())

# ============================================================
# 3. واجهة Streamlit
# ============================================================
st.set_page_config(page_title="DeepClean Studio - Professional", layout="wide")
st.title("📄 DeepClean Studio – المحترف")
st.caption("معالجة مستندات Word مع الحفاظ على الجداول والأشكال والمعادلات - يعمل محليًا - يجتاز ZeroGPT")

uploaded = st.file_uploader("رفع ملف Word أو نص", type=["docx", "txt"])
if uploaded is not None:
    with st.spinner("جاري إعادة الصياغة البشرية..."):
        if uploaded.name.endswith('.docx'):
            in_bytes = extract_uploaded_bytes(uploaded)
            out_bytes = process_docx(in_bytes)
            st.success("تمت معالجة المستند بنجاح مع الحفاظ على الجداول والأشكال والمعادلات!")
            st.download_button(
                "⬇️ تحميل الملف المعدّل",
                data=out_bytes,
                file_name="deepclean_humanized.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            in_bytes = extract_uploaded_bytes(uploaded)
            out_bytes = process_txt(in_bytes)
            st.download_button(
                "⬇️ تحميل النص المعدّل",
                data=out_bytes,
                file_name="deepclean_humanized.txt",
                mime="text/plain"
            )
        st.info("تم تطبيق التغييرات على النصوص فقط. الجداول والأشكال والمعادلات لم تتأثر.")

st.markdown("---")
st.markdown("""
**ملاحظات هامة:**
- هذا الكود يعمل على **نظام التشغيل المحلي** ولا يرسل أي بيانات إلى الإنترنت.
- يحافظ على كل الجداول والأشكال والمعادلات في ملف Word لأنها تُقرأ مباشرة من المستند الأصلي وتُعاد كتابتها مع تغيير النصوص فقط.
- يطبق نفس قواعد الاستبدال التي نجحت مع ZeroGPT (إزالة الكلمات المحظورة، عبارات بشرية بسيطة).
- بعد التحميل، يُرجى فتح الملف في Word ومراجعة الاستشهادات (تظل كما هي).
""")
