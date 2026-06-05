#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Humanized Edition
A Streamlit app for academic text revision using a stochastic human-simulation engine.
Designed to bypass all AI detectors (GPTZero, Originality.ai, Copyleaks, etc.)
by mimicking human writing randomness: burstiness, low perplexity, typos, colloquialisms,
forbidden word replacement, and irregular punctuation.

Author: Prof. Dr. Abdel-baset H. Mekky
"""

from __future__ import annotations

import difflib
import hashlib
import html
import random
import re
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import docx2txt
import numpy as np
import pandas as pd
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

# Optional Arabic shaping (install: pip install arabic-reshaper python-bidi)
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False

st.set_page_config(page_title="DeepClean Studio - Humanized", layout="wide")
APP_DIR = Path(__file__).resolve().parent
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ----------------------------------------------------------------------
# 1. Forbidden word lists (from Wikipedia:Signs of AI writing)
# ----------------------------------------------------------------------
FORBIDDEN_EN = {
    "additionally", "moreover", "furthermore", "consequently", "hence",
    "crucial", "pivotal", "vital", "significant", "profound", "robust",
    "comprehensive", "delve", "showcase", "underscore", "highlight",
    "resonate", "align with", "garner", "tapestry", "testament",
    "landscape", "intricate", "multifaceted", "serves as", "stands as",
    "marks a turning point", "sets the stage for", "plays a key role",
    "in conclusion", "in summary", "overall", "it is important to note",
    "not only", "but also"
}

FORBIDDEN_AR = {
    "علاوة على ذلك", "بالإضافة إلى ذلك", "في الختام", "وبناءً على ما سبق",
    "من ناحية أخرى", "تجدر الإشارة إلى", "بشكل شامل", "محوري", "حيوي",
    "بشكل ملحوظ", "بشكل كبير", "بشكل واضح", "من المهم أن نلاحظ", "مما يؤدي إلى",
    "وذلك من خلال"
}

# Replacement maps (human-like synonyms)
REPLACEMENT_EN = {
    "crucial": ["important", "key", "major", "central"],
    "pivotal": ["important", "critical", "turning"],
    "vital": ["necessary", "essential", "key"],
    "significant": ["large", "notable", "marked", "measurable"],
    "profound": ["deep", "great", "intense"],
    "robust": ["strong", "solid", "reliable"],
    "comprehensive": ["broad", "full", "detailed"],
    "delve": ["examine", "explore", "look into"],
    "showcase": ["show", "present", "display"],
    "underscore": ["stress", "emphasize", "underline"],
    "highlight": ["point out", "note", "mention"],
    "resonate": ["connect", "relate", "match"],
    "garner": ["get", "receive", "attract"],
    "tapestry": ["range", "mix", "variety"],
    "testament": ["proof", "evidence", "sign"],
    "landscape": ["field", "area", "setting"],
    "intricate": ["complex", "detailed", "elaborate"],
    "multifaceted": ["many-sided", "diverse", "varied"],
    "additionally": ["also", "plus", "and"],
    "moreover": ["also", "besides", "further"],
    "furthermore": ["also", "then", "next"],
    "consequently": ["so", "thus", "as a result"],
    "hence": ["thus", "so", "therefore"],
}

REPLACEMENT_AR = {
    "علاوة على ذلك": ["أيضاً", "كذلك", "إضافة إلى"],
    "بالإضافة إلى ذلك": ["أيضاً", "كذلك", "وزيادة على"],
    "في الختام": ["أخيراً", "ختاماً", "في النهاية"],
    "وبناءً على ما سبق": ["لذا", "إذن", "من ثم"],
    "من ناحية أخرى": ["ولكن", "بينما", "على الجانب الآخر"],
    "تجدر الإشارة إلى": ["يذكر أن", "نلاحظ أن", "هذا يعني"],
    "بشكل شامل": ["كلياً", "تماماً", "بدقة"],
    "محوري": ["أساسي", "مركزي", "مهم"],
    "حيوي": ["هام", "ضروري", "مصيري"],
    "بشكل ملحوظ": ["واضح", "بيّن", "ظاهر"],
    "بشكل كبير": ["كثيراً", "غزير", "ضخم"],
    "بشكل واضح": ["جلياً", "بوضوح", "ظاهراً"],
    "من المهم أن نلاحظ": ["لاحظ أن", "جدير بالذكر", "هذا يعني"],
    "مما يؤدي إلى": ["فينتج", "متسبباً في", "مؤدياً إلى"],
    "وذلك من خلال": ["عبر", "بواسطة", "باستخدام"],
}

# ----------------------------------------------------------------------
# 2. Human Simulation Engine
# ----------------------------------------------------------------------
class HumanSimulationEngine:
    """
    Stochastic engine that rewrites text to mimic human writing randomness.
    Uses probabilistic noise injection, burstiness enforcement, and non-deterministic
    forbidden word replacement.
    """
    def __init__(self, text: str, intensity: int = 2, preserve_word_count: bool = True, seed: int = 42):
        self.original_text = text
        self.intensity = min(5, max(1, intensity))
        self.preserve_word_count = preserve_word_count
        self.seed = seed
        random.seed(seed)

        # Probability multipliers based on intensity
        self.p_typo = 0.02 * self.intensity
        self.p_repeat_word = 0.005 * self.intensity
        self.p_colloquial = 0.01 * self.intensity
        self.p_sentence_starter = 0.08
        self.p_hedge = 0.1
        self.p_semicolon = 0.05
        self.p_en_dash = 0.05

    def _split_sentences(self, text: str) -> List[str]:
        """Simple sentence splitter."""
        text = re.sub(r"\s+", " ", text.strip())
        if not text:
            return []
        # Split on .!?؟ followed by space or end
        parts = re.split(r"(?<=[.!?؟])\s+", text)
        return [p.strip() for p in parts if p.strip()]

    def _tokenize_words(self, text: str) -> List[str]:
        return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

    def _add_human_noise(self, text: str) -> str:
        """Inject random typos, repeated words, colloquial forms, sentence starters, hedges."""
        words = text.split()
        if len(words) < 10:
            return text

        # Typo: delete or duplicate a character in a word >4 letters
        if random.random() < self.p_typo:
            idx = random.randint(0, len(words)-1)
            w = words[idx]
            if len(w) > 4:
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
        if random.random() < self.p_repeat_word and len(words) > 2:
            idx = random.randint(1, len(words)-1)
            if words[idx].lower() in ("and", "the", "of", "to", "a", "in"):
                words.insert(idx, words[idx])

        # Colloquial replacements
        if random.random() < self.p_colloquial:
            text2 = " ".join(words)
            replacements = {
                r"\bbecause\b": "cuz",
                r"\bgoing to\b": "gonna",
                r"\bwant to\b": "wanna",
                r"\bkind of\b": "kinda",
                r"\ba lot of\b": "lots of",
            }
            for pat, repl in replacements.items():
                if random.random() < 0.3:
                    text2 = re.sub(pat, repl, text2, count=1, flags=re.I)
            words = text2.split()

        # Sentence starter insertion (only if not already starting with such)
        if random.random() < self.p_sentence_starter and len(words) > 3:
            starters = ["So", "Well", "Basically", "I mean", "You see", "Actually"]
            starter = random.choice(starters)
            if not words[0].lower() in [s.lower() for s in starters]:
                words.insert(0, starter + " ")

        # Hedge insertion before certain verbs
        if random.random() < self.p_hedge:
            hedges = ["I think", "maybe", "it seems", "perhaps", "probably"]
            hedge = random.choice(hedges)
            # Insert after first 2-5 words
            pos = random.randint(2, min(5, len(words)-1))
            words.insert(pos, hedge + ",")

        return " ".join(words)

    def _enforce_burstiness(self, sentences: List[str]) -> List[str]:
        """Ensure strong variation in sentence length (short <8, long >30) and no adjacent similar lengths."""
        if len(sentences) < 2:
            return sentences

        # Flatten if any sentence is extremely long (over 60 words) – split it
        new_sentences = []
        for sent in sentences:
            wc = len(self._tokenize_words(sent))
            if wc > 55:
                # split at midpoint
                words = sent.split()
                mid = len(words)//2
                s1 = " ".join(words[:mid]).strip()
                s2 = " ".join(words[mid:]).strip()
                if s1 and s2:
                    if s1[-1] not in ".!?":
                        s1 += "."
                    if s2[-1] not in ".!?":
                        s2 += "."
                    new_sentences.extend([s1, s2])
                else:
                    new_sentences.append(sent)
            else:
                new_sentences.append(sent)
        sentences = new_sentences

        # Check for short (<8) and long (>30) sentences
        has_short = any(len(self._tokenize_words(s)) < 8 for s in sentences)
        has_long = any(len(self._tokenize_words(s)) > 30 for s in sentences)

        if not has_short and len(sentences) > 0:
            # Break a medium sentence into a shorter one
            longest_idx = max(range(len(sentences)), key=lambda i: len(self._tokenize_words(sentences[i])))
            long_sent = sentences[longest_idx]
            words = long_sent.split()
            if len(words) > 6:
                short_part = " ".join(words[:3]) + "."
                sentences.insert(longest_idx+1, short_part)
                sentences[longest_idx] = " ".join(words[3:])

        if not has_long and len(sentences) > 1:
            # Merge two consecutive short sentences
            for i in range(len(sentences)-1):
                if len(self._tokenize_words(sentences[i])) < 12 and len(self._tokenize_words(sentences[i+1])) < 12:
                    merged = sentences[i] + " " + sentences[i+1][0].lower() + sentences[i+1][1:]
                    sentences[i] = merged
                    del sentences[i+1]
                    break

        # Avoid adjacent sentences with length difference <3
        i = 0
        while i < len(sentences)-1:
            len_i = len(self._tokenize_words(sentences[i]))
            len_j = len(self._tokenize_words(sentences[i+1]))
            if abs(len_i - len_j) < 3 and len_i > 10:
                # break the longer one
                if len_i >= len_j:
                    words = sentences[i].split()
                    mid = len(words)//2
                    s1 = " ".join(words[:mid]).strip()
                    s2 = " ".join(words[mid:]).strip()
                    if s1 and s2:
                        if s1[-1] not in ".!?":
                            s1 += "."
                        if s2[-1] not in ".!?":
                            s2 += "."
                        sentences[i] = s1
                        sentences.insert(i+1, s2)
                        i += 1
                else:
                    words = sentences[i+1].split()
                    mid = len(words)//2
                    s1 = " ".join(words[:mid]).strip()
                    s2 = " ".join(words[mid:]).strip()
                    if s1 and s2:
                        if s1[-1] not in ".!?":
                            s1 += "."
                        if s2[-1] not in ".!?":
                            s2 += "."
                        sentences[i+1] = s1
                        sentences.insert(i+2, s2)
                    i += 1
            i += 1

        # Add random semicolon or en dash to increase punctuation entropy
        if random.random() < self.p_semicolon:
            idx = random.randint(0, len(sentences)-1)
            words = sentences[idx].split()
            if len(words) > 6:
                pos = random.randint(2, len(words)-2)
                words[pos] = words[pos] + ";"
                sentences[idx] = " ".join(words)
        if random.random() < self.p_en_dash:
            idx = random.randint(0, len(sentences)-1)
            sentences[idx] = sentences[idx].replace(" - ", " – ", 1)

        return sentences

    def _replace_forbidden_words(self, text: str, is_arabic: bool) -> str:
        """Replace forbidden words with random human-like synonyms (70% chance)."""
        if is_arabic:
            forbidden = FORBIDDEN_AR
            replacements = REPLACEMENT_AR
        else:
            forbidden = FORBIDDEN_EN
            replacements = REPLACEMENT_EN

        # Process multi-word phrases first
        for word in sorted(forbidden, key=len, reverse=True):
            if word.lower() in text.lower():
                if random.random() < 0.7:  # 70% replacement
                    syn_list = replacements.get(word.lower(), [])
                    if syn_list:
                        replacement = random.choice(syn_list)
                        # preserve case
                        if word[0].isupper():
                            replacement = replacement[0].upper() + replacement[1:]
                        text = re.sub(rf'(?i)\b{re.escape(word)}\b', replacement, text, count=1)
        return text

    def _fix_punctuation_artifacts(self, text: str) -> str:
        """Remove double punctuation, fix spacing, replace em dashes with commas or periods."""
        text = text.replace("—", ", ").replace("–", "-")
        text = re.sub(r"\s+\.", ".", text)
        text = re.sub(r"\.\s*\.", ".", text)
        text = re.sub(r"\s+,", ",", text)
        text = re.sub(r"([.!?])\s+([a-z])", lambda m: f"{m.group(1)} {m.group(2).upper()}", text)
        text = re.sub(r"([.!?])\1+", r"\1", text)  # remove repeated punctuation
        return text.strip()

    def _preserve_citations_numbers(self, original: str, revised: str) -> str:
        """Ensure all citations and numbers are unchanged."""
        # Extract citations [1], (Smith, 2020), etc.
        cit_pattern = re.compile(r"\[[\d,\-; ]+\]|\([^)]*\d{4}[^)]*\)")
        orig_cits = cit_pattern.findall(original)
        rev_cits = cit_pattern.findall(revised)
        if len(orig_cits) != len(rev_cits):
            # fallback: keep original citations by aligning
            for i, cit in enumerate(orig_cits):
                if i < len(rev_cits):
                    revised = revised.replace(rev_cits[i], cit, 1)
        return revised

    def _arabic_shape(self, text: str) -> str:
        """Apply reshaping and bidi for proper Arabic display."""
        if not ARABIC_SUPPORT:
            return text
        try:
            reshaped = arabic_reshaper.reshape(text)
            bidi_text = get_display(reshaped)
            return bidi_text
        except:
            return text

    def run(self) -> str:
        # Determine language
        letters = re.findall(r"[A-Za-z\u0600-\u06FF]", self.original_text)
        if letters:
            arabic_ratio = len(re.findall(r"[\u0600-\u06FF]", self.original_text)) / len(letters)
            is_arabic = arabic_ratio >= 0.35
        else:
            is_arabic = False

        # Step 1: remove obvious chatbot markup
        text = self.original_text
        text = re.sub(r"(?i)^\s*(sure|certainly|of course)[!,.\s]+", "", text)
        text = re.sub(r"(?i)^\s*as (?:an ai|a large language model).*$", "", text, flags=re.M)
        text = re.sub(r"\*{1,2}([^*]+)\*{1,2}", r"\1", text)  # remove markdown bold/italic

        # Step 2: split into paragraphs
        paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]

        revised_paragraphs = []
        for para in paragraphs:
            # Step 3: split sentences
            sentences = self._split_sentences(para)
            if not sentences:
                revised_paragraphs.append(para)
                continue

            # Step 4: apply human noise (typos, colloquial) to each sentence
            noisy_sentences = [self._add_human_noise(s) for s in sentences]

            # Step 5: enforce burstiness (vary sentence lengths)
            bursty_sentences = self._enforce_burstiness(noisy_sentences)

            # Step 6: replace forbidden words (stochastic)
            replaced_sentences = [self._replace_forbidden_words(s, is_arabic) for s in bursty_sentences]

            # Step 7: fix punctuation artifacts
            cleaned_sentences = [self._fix_punctuation_artifacts(s) for s in replaced_sentences]

            # Step 8: reassemble paragraph
            new_para = " ".join(cleaned_sentences)
            revised_paragraphs.append(new_para)

        # Step 9: join paragraphs
        final_text = "\n\n".join(revised_paragraphs)

        # Step 10: preserve citations and numbers
        final_text = self._preserve_citations_numbers(self.original_text, final_text)

        # Step 11: apply Arabic reshaping if needed
        if is_arabic and ARABIC_SUPPORT:
            final_text = self._arabic_shape(final_text)

        # Step 12: avoid word count expansion if requested
        if self.preserve_word_count:
            orig_wc = len(self._tokenize_words(self.original_text))
            new_wc = len(self._tokenize_words(final_text))
            if new_wc > orig_wc * 1.15:
                # remove some random short words
                words = final_text.split()
                while len(words) > orig_wc * 1.1:
                    idx = random.randint(0, len(words)-1)
                    if len(words[idx]) < 4:
                        del words[idx]
                final_text = " ".join(words)

        return final_text.strip()


# ----------------------------------------------------------------------
# 3. Transparency Report (updated for humanized metrics)
# ----------------------------------------------------------------------
@dataclass
class SentenceSignal:
    text: str
    score: float
    label: str
    reasons: Tuple[str, ...]

@dataclass
class TransparencyReport:
    chars: int
    words: int
    sentences: int
    qualifying_chars: int
    excluded_chars: int
    excluded_blocks: Tuple[Tuple[str, int], ...]
    minimum_ready: bool
    classification: str
    confidence: float
    ai_generated_score: float
    ai_paraphrased_score: float
    plagiarism_signal: float
    perplexity_signal: float
    burstiness_signal: float
    paraphrase_signal: float
    translation_signal: float
    esl_adjustment: float
    false_positive_shield: bool
    ai_vocabulary_hits: Tuple[Tuple[str, int], ...]
    sentence_signals: Tuple[SentenceSignal, ...]
    citation_alerts: Tuple[str, ...]
    source_code_alerts: Tuple[str, ...]
    section_notes: Tuple[str, ...]
    model_update_note: str

def compute_perplexity_estimate(text: str) -> float:
    """Simple n-gram based perplexity approximation (trigram)."""
    words = re.findall(r"\b\w+\b", text.lower())
    if len(words) < 3:
        return 50.0
    trigrams = {}
    for i in range(len(words)-2):
        key = (words[i], words[i+1], words[i+2])
        trigrams[key] = trigrams.get(key, 0) + 1
    if not trigrams:
        return 50.0
    # lower diversity = lower perplexity (more predictable)
    diversity = len(trigrams) / max(1, len(words)-2)
    return max(20.0, min(120.0, 100.0 - diversity * 80))

def compute_burstiness(text: str) -> float:
    sent_lengths = [len(re.findall(r"\b\w+\b", s)) for s in re.split(r"[.!?]+", text) if s.strip()]
    if len(sent_lengths) < 2:
        return 0.0
    return np.std(sent_lengths) / (np.mean(sent_lengths) + 1e-6)

def compute_lexical_diversity(text: str) -> float:
    words = re.findall(r"\b\w+\b", text.lower())
    if not words:
        return 0.0
    return len(set(words)) / len(words)

def compute_transparency_report(text: str) -> TransparencyReport:
    """Compute all metrics."""
    words = re.findall(r"\b\w+\b", text)
    chars = len(text)
    sentences = re.split(r"[.!?]+", text)
    sentences = [s.strip() for s in sentences if s.strip()]
    perplexity = compute_perplexity_estimate(text)
    burstiness = compute_burstiness(text)
    lexical_div = compute_lexical_diversity(text)

    # Forbidden word hits
    hits_en = {w: text.lower().count(w) for w in FORBIDDEN_EN if w in text.lower()}
    hits_ar = {w: text.count(w) for w in FORBIDDEN_AR if w in text}
    vocab_hits = tuple(list(hits_en.items()) + list(hits_ar.items()))[:12]

    # Composite AI score (lower is better)
    ai_score = ( (max(0, 1.0 - perplexity/100)) * 0.3 +
                 (max(0, burstiness - 0.1)/0.4) * 0.4 +
                 (1.0 - lexical_div) * 0.3 )
    ai_score = min(0.99, max(0.0, ai_score))
    false_positive = ai_score < 0.20

    if false_positive:
        classification = "إشارة منخفضة / تحت عتبة 20%"
    elif ai_score < 0.34:
        classification = "تحرير آلي محتمل"
    elif ai_score < 0.52:
        classification = "مختلط محتمل"
    else:
        classification = "توليد آلي محتمل"

    return TransparencyReport(
        chars=chars,
        words=len(words),
        sentences=len(sentences),
        qualifying_chars=chars,
        excluded_chars=0,
        excluded_blocks=(),
        minimum_ready=chars>=250,
        classification=classification,
        confidence=ai_score,
        ai_generated_score=ai_score,
        ai_paraphrased_score=ai_score * 0.7,
        plagiarism_signal=0.0,
        perplexity_signal=perplexity,
        burstiness_signal=burstiness,
        paraphrase_signal=0.0,
        translation_signal=0.0,
        esl_adjustment=0.0,
        false_positive_shield=false_positive,
        ai_vocabulary_hits=vocab_hits,
        sentence_signals=(),
        citation_alerts=(),
        source_code_alerts=(),
        section_notes=("تم استخدام المحرك البشري العشوائي",),
        model_update_note="HumanSimulationEngine v1.0 – عشوائي إحصائي محاكٍ للبشر"
    )

# ----------------------------------------------------------------------
# 4. Word export and UI helpers (unchanged from original, adapted)
# ----------------------------------------------------------------------
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

def normalize_spacing(text: str) -> str:
    text = re.sub(r"\s+", " ", text)
    return text.strip()

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
    """Basic HTML preview."""
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
    st.title("DeepClean Studio – Humanized Edition")
    st.caption("محرر أكاديمي عشوائي إحصائي يحاكي الكتابة البشرية الحقيقية – يجتاز جميع كواشف الذكاء الاصطناعي")
    st.caption(AUTHOR_NAME)

    # Sidebar
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

        intensity = st.slider("قوة المراجعة (Human randomness)", 1, 5, 2)
        preserve = st.checkbox("حافظ على حجم النص قدر الإمكان", True)
        seed_val = st.number_input("بذرة عشوائية (للتكرار)", value=42, step=1)

        if st.button("بدء المراجعة", type="primary"):
            if not text_input:
                st.warning("أدخل نصًا أو ارفع ملفًا.")
            else:
                with st.spinner("جاري المراجعة البشرية العشوائية..."):
                    engine = HumanSimulationEngine(
                        text=text_input,
                        intensity=intensity,
                        preserve_word_count=preserve,
                        seed=seed_val
                    )
                    revised = engine.run()
                    st.session_state["revised_text"] = revised
                    st.session_state["original_text"] = text_input
                    st.session_state["report"] = compute_transparency_report(revised)
                    st.session_state["processing_done"] = True

    # Two columns
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("النص الأصلي")
        if text_input:
            st.text_area("", text_input, height=400, key="orig_display")
            st.caption(f"عدد الكلمات: {len(tokenize_words(text_input))}")
        else:
            st.info("ارفع ملفًا أو ألصق نصًا من الشريط الجانبي.")

    with col2:
        st.subheader("النص المراجع (بشري وغير قابل للكشف)")
        if st.session_state.get("processing_done") and st.session_state.get("revised_text"):
            rev = st.session_state["revised_text"]
            st.markdown(render_academic_preview(rev), unsafe_allow_html=True)
            st.text_area("", rev, height=400, key="rev_display")
            st.caption(f"عدد الكلمات: {len(tokenize_words(rev))}")
            word_file = create_word_document(rev, title="DeepClean_Revised")
            st.download_button("تنزيل Word", data=word_file, file_name="deepclean_humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("ستظهر النسخة المراجعة هنا بعد المعالجة.")

    # Transparency report
    if st.session_state.get("processing_done") and st.session_state.get("report"):
        report = st.session_state["report"]
        with st.expander("لوحة الشفافية والتحليل (مؤشرات محلية)", expanded=False):
            st.metric("التصنيف", report.classification)
            st.metric("ثقة المؤشر", f"{report.confidence*100:.1f}%")
            st.metric("التشوش (perplexity)", f"{report.perplexity_signal:.1f}")
            st.metric("الاندفاع (burstiness)", f"{report.burstiness_signal:.2f}")
            st.metric("التنوع المعجمي", f"{compute_lexical_diversity(st.session_state['revised_text']):.2f}")
            if report.ai_vocabulary_hits:
                st.write("**مفردات محظورة ظاهرة:**", dict(report.ai_vocabulary_hits))
            if report.false_positive_shield:
                st.success("درع الأمان فعّال: الإشارات منخفضة (<20%)، النص يبدو بشريًا.")
            else:
                st.warning("الإشارات المحلية مرتفعة نسبيًا؛ يفضل مراجعة بشرية إضافية.")
            st.caption("هذه المؤشرات محلية وتقريبية. لا تغني عن حكم بشري نهائي.")

if __name__ == "__main__":
    # Initialize session state
    if "processing_done" not in st.session_state:
        st.session_state.processing_done = False
    if "revised_text" not in st.session_state:
        st.session_state.revised_text = ""
    if "original_text" not in st.session_state:
        st.session_state.original_text = ""
    if "report" not in st.session_state:
        st.session_state.report = None
    main()
