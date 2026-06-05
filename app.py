#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Final Humanized Edition (Fixed)
Aggressive stochastic engine to bypass all AI detectors (GPTZero, ZeroGPT, Originality.ai)
"""

from __future__ import annotations

import html
import random
import re
from collections import Counter
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import List, Optional, Tuple

import docx2txt
import numpy as np
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(page_title="DeepClean Studio - Final Humanized", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ----------------------------------------------------------------------
# 1. Forbidden words and aggressive replacements
# ----------------------------------------------------------------------
FORBIDDEN_EN = {
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "align with", "garner", "tapestry", "testament",
    "landscape", "intricate", "multifaceted", "serves as", "stands as",
    "marks a turning point", "sets the stage for", "plays a key role",
    "in conclusion", "in summary", "overall", "it is important to note",
    "not only", "but also", "uniquely", "constitute", "trajectories",
    "pronounced", "routinely", "impose", "reducing", "exceeding"
}

REPLACEMENT_MAP = {
    "additionally": "also", "moreover": "also", "furthermore": "then",
    "consequently": "so", "hence": "thus", "crucial": "important",
    "pivotal": "key", "vital": "needed", "significant": "big",
    "profound": "deep", "robust": "strong", "comprehensive": "full",
    "delve": "look into", "showcase": "show", "underscore": "stress",
    "highlight": "point out", "resonate": "connect", "align with": "match",
    "garner": "get", "tapestry": "mix", "testament": "proof",
    "landscape": "field", "intricate": "complex", "multifaceted": "varied",
    "serves as": "is", "stands as": "is", "marks a turning point": "changes",
    "sets the stage for": "leads to", "plays a key role": "helps",
    "in conclusion": "", "in summary": "", "overall": "",
    "it is important to note": "", "not only": "", "but also": "and",
    "uniquely": "", "constitute": "is", "trajectories": "paths",
    "pronounced": "clear", "routinely": "often", "impose": "bring",
    "reducing": "cutting", "exceeding": "above"
}

# ----------------------------------------------------------------------
# 2. Aggressive Human Simulation Engine
# ----------------------------------------------------------------------
class AggressiveHumanEngine:
    def __init__(self, text: str, intensity: int = 3, seed: int = 42):
        self.original_text = text
        self.intensity = min(5, max(1, intensity))
        self.seed = seed
        random.seed(seed)

    def _split_sentences(self, text: str) -> List[str]:
        text = re.sub(r"\s+", " ", text.strip())
        if not text:
            return []
        # Split on .!?؟ followed by space or end, but keep abbreviations
        parts = re.split(r"(?<=[.!?؟])\s+(?=[A-Z\u0600-\u06FF])", text)
        return [p.strip() for p in parts if p.strip()]

    def _tokenize_words(self, text: str) -> List[str]:
        return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

    def _force_sentence_splitting(self, sentences: List[str]) -> List[str]:
        """Split any sentence longer than 20 words into two or three shorter ones."""
        new_sentences = []
        for sent in sentences:
            wc = len(self._tokenize_words(sent))
            if wc > 20:
                # Find a good split point: comma, 'and', 'but', 'or', 'so', 'because'
                split_markers = [r',\s+', r'\s+and\s+', r'\s+but\s+', r'\s+or\s+', r'\s+so\s+', r'\s+because\s+']
                best_pos = -1
                best_marker = None
                for marker in split_markers:
                    matches = list(re.finditer(marker, sent))
                    if matches:
                        # choose marker closest to middle
                        mid = len(sent) // 2
                        for m in matches:
                            if abs(m.start() - mid) < abs(best_pos - mid) or best_pos == -1:
                                best_pos = m.start()
                                best_marker = marker.strip()
                if best_pos != -1:
                    first = sent[:best_pos].strip()
                    second = sent[best_pos + len(best_marker):].strip()
                    if first and second:
                        if first[-1] not in '.!?':
                            first += '.'
                        if second[-1] not in '.!?':
                            second += '.'
                        # Capitalize second
                        second = second[0].upper() + second[1:]
                        new_sentences.extend([first, second])
                        continue
                # If no good split, break at word boundary
                words = sent.split()
                mid = max(5, len(words) // 2)
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

    def _remove_forbidden_words(self, text: str) -> str:
        """100% aggressive replacement of forbidden words."""
        lowered = text.lower()
        for bad, good in REPLACEMENT_MAP.items():
            if bad in lowered:
                # preserve case roughly
                pattern = re.compile(rf'\b{re.escape(bad)}\b', re.I)
                text = pattern.sub(good, text)
        return text

    def _add_human_noise(self, text: str) -> str:
        """Add typos, repeated words, colloquial starters, hedges, and questions."""
        words = text.split()
        if len(words) < 5:
            return text

        # Typo: 5% chance per 50 words
        if random.random() < 0.05 * (self.intensity / 3):
            idx = random.randint(0, len(words)-1)
            w = words[idx]
            if len(w) > 3:
                if random.random() < 0.5:
                    # delete random char
                    pos = random.randint(0, len(w)-1)
                    w = w[:pos] + w[pos+1:]
                else:
                    # duplicate random char
                    pos = random.randint(0, len(w)-1)
                    w = w[:pos] + w[pos] + w[pos:]
                words[idx] = w

        # Repeat a small word
        if random.random() < 0.02 * (self.intensity / 3):
            idx = random.randint(1, len(words)-1)
            if words[idx].lower() in ("the", "and", "of", "to", "a", "in"):
                words.insert(idx, words[idx])

        # Colloquial replacements
        if random.random() < 0.03 * (self.intensity / 3):
            colloq = {
                r'\bbecause\b': 'cuz',
                r'\bgoing to\b': 'gonna',
                r'\bwant to\b': 'wanna',
                r'\bkind of\b': 'kinda',
                r'\ba lot of\b': 'lots of',
                r'\byou are\b': "you're",
            }
            text2 = ' '.join(words)
            for pat, repl in colloq.items():
                if random.random() < 0.3:
                    text2 = re.sub(pat, repl, text2, count=1, flags=re.I)
            words = text2.split()

        # Add sentence starter (So, Well, Look, Basically, I mean) at beginning of first sentence of paragraph
        if random.random() < 0.2 * (self.intensity / 3):
            starters = ['So ', 'Well ', 'Look, ', 'Basically ', 'I mean, ']
            starter = random.choice(starters)
            if not words[0].lower().startswith(('so', 'well', 'look', 'basically', 'i mean')):
                words.insert(0, starter)

        # Add hedge (I think, maybe, it seems) after first few words
        if random.random() < 0.15 * (self.intensity / 3):
            hedges = ['I think ', 'maybe ', 'it seems ', 'perhaps ', 'probably ']
            hedge = random.choice(hedges)
            pos = random.randint(1, min(5, len(words)-1))
            words.insert(pos, hedge.rstrip())

        # Add question tag at end of a random sentence (but not too often)
        if random.random() < 0.1:
            if words[-1][-1] in '.!?':
                words[-1] = words[-1][:-1] + ', right?'
            else:
                words[-1] = words[-1] + ', right?'

        return ' '.join(words)

    def _enforce_burstiness(self, sentences: List[str]) -> List[str]:
        """Ensure at least one very short sentence (<8 words) and one long (>25 words) per paragraph."""
        if len(sentences) < 2:
            return sentences

        # Count short and long
        short_count = sum(1 for s in sentences if len(self._tokenize_words(s)) < 8)
        long_count = sum(1 for s in sentences if len(self._tokenize_words(s)) > 25)

        # If no short sentence, break a longer sentence into a very short one
        if short_count == 0 and len(sentences) > 0:
            # find longest sentence
            longest_idx = max(range(len(sentences)), key=lambda i: len(self._tokenize_words(sentences[i])))
            long_sent = sentences[longest_idx]
            words = long_sent.split()
            if len(words) > 5:
                short_part = ' '.join(words[:3]) + '.'
                long_rest = ' '.join(words[3:])
                if long_rest and long_rest[-1] not in '.!?':
                    long_rest += '.'
                sentences[longest_idx] = long_rest
                sentences.insert(longest_idx + 1, short_part)
                short_count += 1

        # If no long sentence, merge two short sentences
        if long_count == 0 and len(sentences) > 1:
            for i in range(len(sentences)-1):
                if len(self._tokenize_words(sentences[i])) < 12 and len(self._tokenize_words(sentences[i+1])) < 12:
                    merged = sentences[i] + ' ' + sentences[i+1][0].lower() + sentences[i+1][1:]
                    sentences[i] = merged
                    del sentences[i+1]
                    break

        # Ensure no adjacent sentences have similar length (difference <3)
        i = 0
        while i < len(sentences)-1:
            len_i = len(self._tokenize_words(sentences[i]))
            len_j = len(self._tokenize_words(sentences[i+1]))
            if abs(len_i - len_j) < 3 and len_i > 8:
                # split the longer one
                if len_i >= len_j:
                    words = sentences[i].split()
                    if len(words) > 4:
                        mid = len(words)//2
                        s1 = ' '.join(words[:mid]).strip()
                        s2 = ' '.join(words[mid:]).strip()
                        if s1 and s2:
                            if s1[-1] not in '.!?':
                                s1 += '.'
                            if s2[-1] not in '.!?':
                                s2 += '.'
                            sentences[i] = s1
                            sentences.insert(i+1, s2)
                            i += 1
                else:
                    words = sentences[i+1].split()
                    if len(words) > 4:
                        mid = len(words)//2
                        s1 = ' '.join(words[:mid]).strip()
                        s2 = ' '.join(words[mid:]).strip()
                        if s1 and s2:
                            if s1[-1] not in '.!?':
                                s1 += '.'
                            if s2[-1] not in '.!?':
                                s2 += '.'
                            sentences[i+1] = s1
                            sentences.insert(i+2, s2)
                            i += 1
            i += 1
        return sentences

    def _preserve_citations(self, original: str, revised: str) -> str:
        """Restore original citations exactly."""
        cit_pattern = re.compile(r'\[\d+(?:[-,;]\s*\d+)*\]|\([^)]*\d{4}[^)]*\)')
        orig_cits = cit_pattern.findall(original)
        rev_cits = cit_pattern.findall(revised)
        if len(orig_cits) != len(rev_cits):
            # fallback: replace sequentially
            for i, oc in enumerate(orig_cits):
                if i < len(rev_cits):
                    revised = revised.replace(rev_cits[i], oc, 1)
        return revised

    def _fix_grammar_artifacts(self, text: str) -> str:
        """Clean up double punctuation, spacing, and em dashes. Fixed regex group error."""
        text = text.replace('—', ', ').replace('–', '-')
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\.\s*\.', '.', text)
        text = re.sub(r'\s+,', ',', text)
        text = re.sub(r'([.!?])\s+([a-z])', lambda m: f'{m.group(1)} {m.group(2).upper()}', text)
        text = re.sub(r'([.!?])\1+', r'\1', text)
        # Fix common broken patterns like "above, depth" -> "above, and depth" - corrected groups
        # The pattern has two capturing groups: (\w+), (\w+). Use \1 and \2 only.
        text = re.sub(r'(\w+), (\w+) (?:exceeding|above|below)', r'\1, and \2', text)
        return text.strip()

    def run(self) -> str:
        text = self.original_text
        # Step 1: remove obvious chatbot markup
        text = re.sub(r'(?i)^\s*(sure|certainly|of course)[!,.\s]+', '', text)
        text = re.sub(r'(?i)^\s*as (?:an ai|a large language model).*$', '', text, flags=re.M)
        text = re.sub(r'\*{1,2}([^*]+)\*{1,2}', r'\1', text)

        # Step 2: split into paragraphs
        paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
        revised_paragraphs = []

        for para in paragraphs:
            # Step 3: split sentences
            sentences = self._split_sentences(para)
            if not sentences:
                revised_paragraphs.append(para)
                continue

            # Step 4: aggressively split long sentences (>20 words)
            sentences = self._force_sentence_splitting(sentences)

            # Step 5: remove forbidden words completely
            sentences = [self._remove_forbidden_words(s) for s in sentences]

            # Step 6: add human noise (typos, colloquial, starters)
            sentences = [self._add_human_noise(s) for s in sentences]

            # Step 7: enforce burstiness (short + long sentences)
            sentences = self._enforce_burstiness(sentences)

            # Step 8: fix punctuation artifacts
            sentences = [self._fix_grammar_artifacts(s) for s in sentences]

            # Reassemble paragraph
            new_para = ' '.join(sentences)
            revised_paragraphs.append(new_para)

        final_text = '\n\n'.join(revised_paragraphs)

        # Step 9: restore citations
        final_text = self._preserve_citations(self.original_text, final_text)

        # Step 10: ensure no em dashes or stray markdown
        final_text = final_text.replace('**', '').replace('__', '')
        final_text = re.sub(r'https?://\S+', '', final_text)  # remove raw URLs if any

        return final_text.strip()


# ----------------------------------------------------------------------
# 3. Simplified Transparency Report (just for UI)
# ----------------------------------------------------------------------
@dataclass
class SimpleReport:
    classification: str
    confidence: float
    false_positive_shield: bool
    ai_vocabulary_hits: dict

def quick_report(text: str) -> SimpleReport:
    words = re.findall(r'\b\w+\b', text.lower())
    forbidden_hits = {w: text.lower().count(w) for w in FORBIDDEN_EN if w in text.lower()}
    score = len(forbidden_hits) / (len(words)+1)
    confidence = min(0.99, score * 2)
    false_shield = confidence < 0.20
    if false_shield:
        classification = "إشارة منخفضة / تحت عتبة 20%"
    elif confidence < 0.34:
        classification = "تحرير آلي محتمل"
    elif confidence < 0.52:
        classification = "مختلط محتمل"
    else:
        classification = "توليد آلي محتمل"
    return SimpleReport(classification, confidence, false_shield, forbidden_hits)


# ----------------------------------------------------------------------
# 4. UI Helpers (extract, word export)
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
    doc.core_properties.title = title or "DeepClean Humanized Revised Manuscript"
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

def render_academic_preview(text: str) -> str:
    parts = ["<div class='academic-preview'>"]
    for line in text.split("\n"):
        if line.strip():
            parts.append(f"<p>{html.escape(line)}</p>")
        else:
            parts.append("<p><br></p>")
    parts.append("</div>")
    return "".join(parts)


# ----------------------------------------------------------------------
# 5. Streamlit UI
# ----------------------------------------------------------------------
def main():
    st.title("DeepClean Studio – Final Humanized Edition (Fixed)")
    st.caption("محرر عشوائي عدواني يحاكي الأخطاء البشرية – يجتاز GPTZero و ZeroGPT و Originality.ai")
    st.caption(AUTHOR_NAME)

    with st.sidebar:
        st.header("الإعدادات")
        input_option = st.radio("مصدر النص", ("رفع ملف", "لصق نص"), key="src_radio")
        text_input = ""
        if input_option == "رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                text_input = extract_uploaded_text(uploaded)
        else:
            text_input = st.text_area("ألصق النص الأكاديمي هنا", height=220)

        intensity = st.slider("قوة المراجعة (عدوانية)", 1, 5, 3)
        seed_val = st.number_input("بذرة عشوائية", value=42, step=1)

        if st.button("بدء المراجعة", type="primary"):
            if not text_input:
                st.warning("أدخل نصًا أو ارفع ملفًا.")
            else:
                with st.spinner("جاري المراجعة البشرية العدوانية..."):
                    engine = AggressiveHumanEngine(text=text_input, intensity=intensity, seed=seed_val)
                    revised = engine.run()
                    st.session_state["revised_text"] = revised
                    st.session_state["original_text"] = text_input
                    st.session_state["report"] = quick_report(revised)
                    st.session_state["processing_done"] = True

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("النص الأصلي")
        if text_input:
            st.text_area("", text_input, height=400, key="orig_display")
            st.caption(f"كلمات: {len(tokenize_words(text_input))}")
        else:
            st.info("ارفع ملفًا أو ألصق نصًا.")

    with col2:
        st.subheader("النص المراجع (بشري وغير قابل للكشف)")
        if st.session_state.get("processing_done") and st.session_state.get("revised_text"):
            rev = st.session_state["revised_text"]
            st.markdown(render_academic_preview(rev), unsafe_allow_html=True)
            st.text_area("", rev, height=400, key="rev_display")
            st.caption(f"كلمات: {len(tokenize_words(rev))}")
            word_file = create_word_document(rev, title="DeepClean_Humanized")
            st.download_button("تنزيل Word", data=word_file, file_name="deepclean_humanized_final.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("ستظهر النسخة المراجعة هنا بعد المعالجة.")

    if st.session_state.get("processing_done") and st.session_state.get("report"):
        rep = st.session_state["report"]
        with st.expander("نتيجة الفحص المحلي (تقديري)", expanded=False):
            st.metric("التصنيف", rep.classification)
            st.metric("درجة الثقة", f"{rep.confidence*100:.1f}%")
            if rep.false_positive_shield:
                st.success("✓ درع الأمان فعال: الإشارات منخفضة، النص يبدو بشريًا.")
            else:
                st.warning("⚠ الإشارات مرتفعة نسبيًا، قد يحتاج النص إلى مراجعة إضافية.")
            if rep.ai_vocabulary_hits:
                st.write("**كلمات محظورة متبقية:**", rep.ai_vocabulary_hits)

if __name__ == "__main__":
    if "processing_done" not in st.session_state:
        st.session_state.processing_done = False
    if "revised_text" not in st.session_state:
        st.session_state.revised_text = ""
    if "original_text" not in st.session_state:
        st.session_state.original_text = ""
    if "report" not in st.session_state:
        st.session_state.report = None
    main()
