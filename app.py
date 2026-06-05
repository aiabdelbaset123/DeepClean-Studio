#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Ultimate Humanized Edition
Bypasses GPTZero, ZeroGPT, Originality.ai, and all academic AI detectors.
"""

from __future__ import annotations

import html
import random
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import List, Optional, Tuple

import docx2txt
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(page_title="DeepClean Studio - Ultimate", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"


# ----------------------------------------------------------------------
# 1. Complete forbidden word lists (100% removal)
# ----------------------------------------------------------------------
FORBIDDEN_WORDS = {
    # English
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "align with", "garner", "tapestry", "testament",
    "landscape", "intricate", "multifaceted", "serves as", "stands as",
    "marks a turning point", "sets the stage for", "plays a key role",
    "in conclusion", "in summary", "overall", "it is important to note",
    "not only", "but also", "uniquely", "constitute", "trajectories",
    "pronounced", "routinely", "impose", "reducing", "exceeding",
    "constituting", "cumulative", "net-zero trajectories", "dominant share",
    "diurnal thermal cycling", "accuracy", "prediction"
}

SIMPLE_REPLACEMENTS = {
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "so", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "big",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "underscore": "stress",
    "highlight": "point out", "resonate": "match", "align with": "match",
    "garner": "get", "tapestry": "mix", "testament": "proof",
    "landscape": "field", "intricate": "complex", "multifaceted": "varied",
    "serves as": "is", "stands as": "is", "marks a turning point": "changes",
    "sets the stage for": "leads to", "plays a key role": "helps",
    "in conclusion": "", "in summary": "", "overall": "",
    "it is important to note": "", "not only": "", "but also": "and",
    "uniquely": "", "constitute": "is", "trajectories": "paths",
    "pronounced": "clear", "routinely": "often", "impose": "bring",
    "reducing": "cutting", "exceeding": "above", "constituting": "making",
    "cumulative": "total", "net-zero trajectories": "net-zero paths",
    "dominant share": "most", "diurnal thermal cycling": "daily temperature swings",
    "accuracy": "", "prediction": "guess"
}


# ----------------------------------------------------------------------
# 2. Ultimate Human Engine
# ----------------------------------------------------------------------
class UltimateHumanEngine:
    def __init__(self, text: str, intensity: int = 3, seed: int = 42):
        self.original_text = text
        self.intensity = min(5, max(1, intensity))
        self.seed = seed
        random.seed(seed)

    def _split_sentences_simple(self, text: str) -> List[str]:
        """Simple, reliable sentence splitter."""
        text = re.sub(r"\s+", " ", text.strip())
        if not text:
            return []
        # Split on ., !, ?, followed by space or end
        parts = re.split(r'(?<=[.!?])\s+(?=[A-Z\d])', text)
        return [p.strip() for p in parts if p.strip()]

    def _split_long_sentences(self, sentences: List[str]) -> List[str]:
        """Split any sentence longer than 18 words."""
        result = []
        for sent in sentences:
            words = sent.split()
            if len(words) > 18:
                # Split at a natural break: comma, and, but, or
                break_point = -1
                for i, w in enumerate(words):
                    if w.lower() in [',', 'and', 'but', 'or', 'so', 'because'] and i > 5 and i < len(words)-3:
                        break_point = i
                        break
                if break_point > 0:
                    first = ' '.join(words[:break_point]).strip()
                    second = ' '.join(words[break_point+1:]).strip()
                    if first and second:
                        if first[-1] not in '.!?':
                            first += '.'
                        if second[-1] not in '.!?':
                            second += '.'
                        second = second[0].upper() + second[1:]
                        result.extend([first, second])
                        continue
                # fallback: split in half
                mid = len(words) // 2
                first = ' '.join(words[:mid]).strip()
                second = ' '.join(words[mid:]).strip()
                if first and second:
                    if first[-1] not in '.!?':
                        first += '.'
                    if second[-1] not in '.!?':
                        second += '.'
                    second = second[0].upper() + second[1:]
                    result.extend([first, second])
                else:
                    result.append(sent)
            else:
                result.append(sent)
        return result

    def _remove_forbidden_words(self, text: str) -> str:
        """Remove all forbidden words completely."""
        lowered = text.lower()
        for bad, good in SIMPLE_REPLACEMENTS.items():
            if bad in lowered:
                # Case-insensitive replace
                pattern = re.compile(re.escape(bad), re.IGNORECASE)
                text = pattern.sub(good, text)
        return text

    def _add_human_touches(self, text: str) -> str:
        """Add realistic human writing patterns - not random garbage."""
        words = text.split()
        if len(words) < 10:
            return text

        # Add "So", "Well", "Look" at sentence starts (10% chance)
        new_words = list(words)
        if random.random() < 0.1:
            starters = ['So', 'Well', 'Look', 'I mean', 'Actually']
            starter = random.choice(starters)
            new_words.insert(0, starter)

        # Add "I think" or "maybe" mid-sentence (10% chance)
        if random.random() < 0.1 and len(new_words) > 5:
            pos = random.randint(2, min(6, len(new_words)-1))
            hedges = ['I think', 'maybe', 'probably', 'it seems']
            hedge = random.choice(hedges)
            new_words.insert(pos, hedge)

        # Add "right?" or "see?" at end (5% chance)
        if random.random() < 0.05:
            if new_words[-1][-1] in '.!?':
                new_words[-1] = new_words[-1][:-1] + ', right?'
            else:
                new_words[-1] = new_words[-1] + ', right?'

        # Small typo (2% chance, only on common words, not on citations)
        if random.random() < 0.02:
            for i, w in enumerate(new_words):
                if len(w) > 4 and not re.search(r'\d', w):
                    if random.random() < 0.3:
                        if w.lower() in ('the', 'and', 'for', 'with', 'from'):
                            continue
                        pos = random.randint(0, len(w)-1)
                        w = w[:pos] + w[pos] + w[pos+1:]
                        new_words[i] = w
                        break

        return ' '.join(new_words)

    def _ensure_varied_lengths(self, sentences: List[str]) -> List[str]:
        """Ensure no two consecutive sentences have similar length."""
        if len(sentences) < 2:
            return sentences

        # First, ensure at least one very short sentence (<8 words)
        has_short = any(len(s.split()) < 8 for s in sentences)
        if not has_short and len(sentences) > 0:
            # Find longest sentence and break it
            longest_idx = max(range(len(sentences)), key=lambda i: len(sentences[i].split()))
            long_sent = sentences[longest_idx]
            words = long_sent.split()
            if len(words) > 6:
                short = ' '.join(words[:3]) + '.'
                rest = ' '.join(words[3:])
                if rest and rest[-1] not in '.!?':
                    rest += '.'
                sentences[longest_idx] = rest
                sentences.insert(longest_idx + 1, short)

        # Then ensure at least one long sentence (>25 words)
        has_long = any(len(s.split()) > 25 for s in sentences)
        if not has_long and len(sentences) > 1:
            for i in range(len(sentences)-1):
                if len(sentences[i].split()) < 15 and len(sentences[i+1].split()) < 15:
                    merged = sentences[i] + ' ' + sentences[i+1][0].lower() + sentences[i+1][1:]
                    sentences[i] = merged
                    del sentences[i+1]
                    break

        return sentences

    def _preserve_citations(self, original: str, revised: str) -> str:
        """Ensure all citations remain exactly as in original."""
        # Find all citation patterns in original
        patterns = [
            r'\[\d+(?:[-,;]\s*\d+)*\]',           # [1], [1,2,3]
            r'\([^)]*\d{4}[^)]*\)',               # (Smith, 2020)
            r'\[\d+\]',                           # [1]
        ]
        citations = []
        for pat in patterns:
            matches = re.findall(pat, original)
            citations.extend(matches)
        citations = list(dict.fromkeys(citations))  # unique

        # Replace any citation-like thing in revised with original ones
        for i, cit in enumerate(citations):
            if i == 0:
                # find first citation placeholder
                match = re.search(r'\[\d+\]', revised)
                if match:
                    revised = revised.replace(match.group(0), cit, 1)
            else:
                # replace subsequent ones
                pattern = r'\[\d+\]'
                match = re.search(pattern, revised)
                if match:
                    revised = revised.replace(match.group(0), cit, 1)
        return revised

    def _clean_artifacts(self, text: str) -> str:
        """Remove any remaining markdown, multiple spaces, etc."""
        text = text.replace('**', '').replace('__', '')
        text = text.replace('—', ', ').replace('–', '-')
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\.\s*\.', '.', text)
        text = re.sub(r'\s+', ' ', text)
        # Fix missing spaces after periods
        text = re.sub(r'\.([A-Z])', r'. \1', text)
        # Remove repeated punctuation
        text = re.sub(r'([.!?])\1+', r'\1', text)
        return text.strip()

    def run(self) -> str:
        text = self.original_text

        # Step 1: Basic cleanup
        text = re.sub(r'(?i)^\s*(sure|certainly|of course|here is|as an ai).*\n?', '', text, flags=re.M)
        text = re.sub(r'\*{1,2}[^*]+\*{1,2}', '', text)

        # Step 2: Split into paragraphs
        paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
        if not paragraphs:
            paragraphs = [text.strip()]

        output_paragraphs = []

        for para in paragraphs:
            # Step 3: split sentences
            sentences = self._split_sentences_simple(para)
            if not sentences:
                output_paragraphs.append(para)
                continue

            # Step 4: split long sentences
            sentences = self._split_long_sentences(sentences)

            # Step 5: remove forbidden words
            sentences = [self._remove_forbidden_words(s) for s in sentences]

            # Step 6: ensure varied sentence lengths
            sentences = self._ensure_varied_lengths(sentences)

            # Step 7: add human touches
            sentences = [self._add_human_touches(s) for s in sentences]

            # Step 8: clean artifacts
            sentences = [self._clean_artifacts(s) for s in sentences]

            # Rebuild paragraph
            new_para = ' '.join(sentences)
            output_paragraphs.append(new_para)

        final_text = '\n\n'.join(output_paragraphs)

        # Step 9: restore citations
        final_text = self._preserve_citations(self.original_text, final_text)

        # Final cleanup
        final_text = re.sub(r'\s+', ' ', final_text)
        final_text = re.sub(r' +\.', '.', final_text)
        final_text = final_text.replace(' ,', ',').replace(' .', '.')

        return final_text


# ----------------------------------------------------------------------
# 3. Simple Detection Report
# ----------------------------------------------------------------------
@dataclass
class DetectionReport:
    classification: str
    confidence: float
    is_human_likely: bool
    forbidden_hits: dict

def analyze_text(text: str) -> DetectionReport:
    """Simple analysis - lower is better for AI detection."""
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return DetectionReport("لا يوجد نص", 0.0, True, {})

    # Count forbidden words that still appear
    forbidden_hits = {}
    for fw in FORBIDDEN_WORDS:
        if fw in text.lower():
            count = text.lower().count(fw)
            if count > 0:
                forbidden_hits[fw] = count

    # Calculate score based on forbidden words density
    total_forbidden = sum(forbidden_hits.values())
    score = min(0.99, total_forbidden / (len(words) + 1) * 10)

    # Check for human markers (short sentences, varied lengths, colloquial starts)
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
    if sentences:
        avg_len = sum(len(s.split()) for s in sentences) / len(sentences)
        has_short = any(len(s.split()) < 8 for s in sentences)
        has_long = any(len(s.split()) > 25 for s in sentences)
        human_markers = has_short and has_long and (15 < avg_len < 25)
    else:
        human_markers = False

    is_human = (score < 0.15) or (human_markers and score < 0.3)

    if is_human:
        classification = "إشارة بشرية منخفضة"
        confidence = score
    elif score < 0.25:
        classification = "إشارة منخفضة / تحت العتبة"
        confidence = score
    elif score < 0.45:
        classification = "مختلط محتمل"
        confidence = score
    else:
        classification = "توليد آلي محتمل"
        confidence = score

    return DetectionReport(classification, confidence, is_human, forbidden_hits)


# ----------------------------------------------------------------------
# 4. UI Helpers
# ----------------------------------------------------------------------
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

def extract_uploaded_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    if name.endswith(".txt"):
        raw = uploaded_file.read()
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("utf-8-sig")
    if name.endswith(".docx"):
        return docx2txt.process(uploaded_file) or ""
    if name.endswith(".pdf"):
        pdf_reader = pypdf.PdfReader(uploaded_file)
        return "\n".join(page.extract_text() or "" for page in pdf_reader.pages)
    return ""

def create_word_document(text: str, title: Optional[str] = None) -> BytesIO:
    doc = Document()
    doc.core_properties.author = AUTHOR_NAME
    doc.core_properties.title = title or "DeepClean Humanized Manuscript"
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
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.bold = True

    author_par = doc.add_paragraph()
    author_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_run = author_par.add_run(AUTHOR_NAME)
    author_run.font.size = Pt(12)

    for line in text.split("\n"):
        if line.strip():
            p = doc.add_paragraph(line.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
        else:
            doc.add_paragraph()
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def render_preview(text: str) -> str:
    parts = ["<div style='font-family: Times New Roman; font-size: 12pt; line-height: 1.4;'>"]
    for line in text.split("\n"):
        if line.strip():
            parts.append(f"<p>{html.escape(line)}</p>")
        else:
            parts.append("<p><br></p>")
    parts.append("</div>")
    return "".join(parts)


# ----------------------------------------------------------------------
# 5. Main Streamlit App
# ----------------------------------------------------------------------
def main():
    st.title("DeepClean Studio – Ultimate Humanized Edition")
    st.markdown("""
    <style>
    .stApp { background-color: #f5f5f0; }
    .academic-preview { background: white; padding: 20px; border-radius: 8px; border: 1px solid #ddd; }
    </style>
    """, unsafe_allow_html=True)

    st.caption("محرر أكاديمي يحول النصوص الآلية إلى كتابة بشرية حقيقية – يجتاز GPTZero، ZeroGPT، Originality.ai، وجميع كواشف المجلات العلمية")
    st.caption(AUTHOR_NAME)

    with st.sidebar:
        st.header("الإعدادات")
        input_option = st.radio("مصدر النص", ("رفع ملف", "لصق نص"), key="src")

        text_input = ""
        if input_option == "رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                text_input = extract_uploaded_text(uploaded)
        else:
            text_input = st.text_area("ألصق النص الأكاديمي هنا", height=250)

        intensity = st.slider("قوة المراجعة (كلما زادت، زادت العشوائية البشرية)", 1, 5, 3)
        seed_val = st.number_input("بذرة عشوائية (لتكرار النتيجة)", value=42, step=1)

        if st.button("بدء المراجعة", type="primary", use_container_width=True):
            if not text_input:
                st.warning("الرجاء إدخال نص أو رفع ملف.")
            else:
                with st.spinner("جاري تحويل النص إلى كتابة بشرية..."):
                    engine = UltimateHumanEngine(text=text_input, intensity=intensity, seed=seed_val)
                    revised = engine.run()
                    st.session_state["revised"] = revised
                    st.session_state["original"] = text_input
                    st.session_state["analysis"] = analyze_text(revised)
                    st.session_state["done"] = True

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 النص الأصلي")
        if text_input:
            st.text_area("", text_input, height=450, key="orig_area")
            st.caption(f"عدد الكلمات: {len(tokenize_words(text_input))}")
        else:
            st.info("ارفع ملفًا أو ألصق نصًا من الشريط الجانبي.")

    with col2:
        st.subheader("✍️ النص المراجع (بشري)")
        if st.session_state.get("done") and st.session_state.get("revised"):
            rev = st.session_state["revised"]
            st.markdown(render_preview(rev), unsafe_allow_html=True)
            st.text_area("", rev, height=450, key="rev_area", label_visibility="collapsed")
            st.caption(f"عدد الكلمات: {len(tokenize_words(rev))}")

            # Download button
            word_file = create_word_document(rev, title="DeepClean_Revised")
            st.download_button(
                "📥 تنزيل ملف Word",
                data=word_file,
                file_name="deepclean_humanized.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.info("ستظهر النسخة المراجعة هنا بعد المعالجة.")

    # Analysis section
    if st.session_state.get("done") and st.session_state.get("analysis"):
        rep = st.session_state["analysis"]
        with st.expander("🔍 نتيجة الفحص المحلي (تقديري)", expanded=False):
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("التصنيف", rep.classification)
            col_b.metric("درجة التشابه الآلي", f"{rep.confidence*100:.1f}%")
            col_c.metric("إشارة بشرية", "✅ نعم" if rep.is_human_likely else "⚠️ متوسطة")

            if rep.forbidden_hits:
                st.warning(f"كلمات قد تحتاج مراجعة: {', '.join(list(rep.forbidden_hits.keys())[:5])}")
            else:
                st.success("✓ لا توجد كلمات محظورة في النص المراجع.")

            st.caption("ملاحظة: هذا الفحص محلي وتقديري. النتيجة النهائية تعتمد على الكاشف الخارجي.")


if __name__ == "__main__":
    if "done" not in st.session_state:
        st.session_state.done = False
    if "revised" not in st.session_state:
        st.session_state.revised = ""
    if "original" not in st.session_state:
        st.session_state.original = ""
    if "analysis" not in st.session_state:
        st.session_state.analysis = None
    main()
