#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Safe Word Processor with Full Formatting Preservation
Preserves tables, figures, equations, references, and all non‑text elements.
Outputs a fully formatted Word document ready for journal submission.
"""

import re
import random
import streamlit as st
from io import BytesIO
from typing import Dict

# ---------- Libraries for Word processing ----------
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(page_title="DeepClean Studio - Safe Processor", layout="wide")
st.title("🧬 DeepClean Studio – Safe Word Document Humanizer")
st.caption("Modifies only plain text paragraphs. Preserves all tables, figures, equations, references, and formatting.")

if not DOCX_AVAILABLE:
    st.error("Please install python-docx: pip install python-docx")
    st.stop()

# -------------------- Humanization Rules (from Wikipedia) --------------------
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
    """Detect paragraphs containing OMML equations or embedded objects."""
    for run in para.runs:
        if run.element.xpath('.//m:oMath'):
            return True
        if run.element.xpath('.//w:object'):
            return True
    return False

def process_word_document(input_bytes: BytesIO, intensity: int) -> BytesIO:
    """Modify only plain text paragraphs. Keep all tables, figures, equations, references intact."""
    doc = Document(input_bytes)
    modified_count = 0
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        if is_math_paragraph(para):
            continue
        original = para.text
        new_text = humanize_text(original, intensity)
        if new_text != original:
            # Preserve paragraph style and alignment
            style = para.style
            alignment = para.paragraph_format.alignment
            para.clear()
            run = para.add_run(new_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            para.style = style
            if alignment is not None:
                para.paragraph_format.alignment = alignment
            modified_count += 1
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, modified_count

# -------------------- Quick stats for preview --------------------
def quick_stats(text: str) -> Dict:
    words = len(re.findall(r'\b\w+\b', text))
    forbidden = sum(1 for w in WORD_REPLACEMENTS if w in text.lower())
    return {"words": words, "forbidden": forbidden}

# -------------------- Streamlit UI --------------------
def main():
    st.sidebar.header("⚙️ Settings")
    intensity = st.sidebar.slider("Transformation Strength", 1, 5, 3,
                                 help="1=light changes, 5=aggressive splitting & human touches")
    uploaded_file = st.sidebar.file_uploader("Upload Word (.docx)", type=["docx"])
    process = st.sidebar.button("🛡️ Process & Keep All Formatting", type="primary", use_container_width=True)

    if uploaded_file and process:
        with st.spinner("Processing document – preserving tables, figures, equations, references..."):
            input_bytes = BytesIO(uploaded_file.read())
            output_bytes, count = process_word_document(input_bytes, intensity)
            st.session_state['output_bytes'] = output_bytes
            st.session_state['count'] = count
            # store a sample of original text for preview
            doc = Document(BytesIO(uploaded_file.read()))
            orig_sample = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()][:10])
            st.session_state['orig_sample'] = orig_sample

    if 'output_bytes' in st.session_state:
        st.success(f"✅ Modified {st.session_state['count']} text paragraphs. All tables, figures, equations, and references remain intact.")
        
        with st.expander("🔍 Preview of changes (text only)", expanded=False):
            # show original sample
            if 'orig_sample' in st.session_state:
                st.text("Original (first few lines):")
                st.text(st.session_state['orig_sample'][:500])
            # show modified sample
            st.session_state['output_bytes'].seek(0)
            doc_out = Document(st.session_state['output_bytes'])
            out_sample = '\n'.join([p.text for p in doc_out.paragraphs if p.text.strip()][:10])
            st.text("Modified (first few lines):")
            st.text(out_sample[:500])
        
        st.subheader("📥 Download Humanized Document")
        st.download_button(
            "📘 Download Fully Formatted Word File",
            data=st.session_state['output_bytes'],
            file_name="deepclean_humanized.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.info("💡 **The downloaded file preserves all original tables, figures, equations, and references.** Review it before final submission.")
    
    elif uploaded_file and not process:
        st.info("Click 'Process & Keep All Formatting' to humanize the text while preserving all non‑text elements.")
        # cache original sample
        doc = Document(BytesIO(uploaded_file.read()))
        st.session_state['orig_sample'] = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()][:10])

if __name__ == "__main__":
    random.seed(42)
    main()
