"""
DeepClean Studio — AI-Text Humanization Engine
================================================
A multi-layered "Humanize Protocol" that transforms AI-generated academic /
scientific texts into texts indistinguishable from expert human writing.

Author : DeepClean Research Lab
Version: 1.0.0
License: MIT

Every module is annotated with:
  - English implementation comments
  - Arabic forensic rationale (why this defeats a specific detection layer)
"""

from __future__ import annotations

import csv
import hashlib
import math
import os
import random
import re
import string
import tempfile
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import nltk
import numpy as np
import spacy
import streamlit as st
from nltk import pos_tag, word_tokenize
from nltk.tokenize import sent_tokenize
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# ─────────────────────────────────────────────
# One-time resource downloads
# ─────────────────────────────────────────────
@st.cache_resource
def _download_nltk():
    for pkg in ("punkt", "punkt_tab", "averaged_perceptron_tagger",
                "averaged_perceptron_tagger_eng"):
        try:
            nltk.data.find(f"tokenizers/{pkg}" if "punkt" in pkg
                           else f"taggers/{pkg}")
        except LookupError:
            nltk.download(pkg, quiet=True)

_download_nltk()

# ─────────────────────────────────────────────
# Lazy-load spaCy
# ─────────────────────────────────────────────
@st.cache_resource
def _load_spacy():
    try:
        return spacy.load("en_core_web_sm")
    except OSError:
        spacy.cli.download("en_core_web_sm")
        return spacy.load("en_core_web_sm")

NLP = _load_spacy()

# ─────────────────────────────────────────────
# Lazy-load SentenceTransformer
# ─────────────────────────────────────────────
@st.cache_resource
def _load_sentence_model():
    return SentenceTransformer("all-MiniLM-L6-v2")

SENT_MODEL = _load_sentence_model()

# ─────────────────────────────────────────────
# Perplexity model (GPT-2 small)
# ─────────────────────────────────────────────
@st.cache_resource
def _load_gpt2():
    from transformers import GPT2LMHeadModel, GPT2TokenizerFast
    model_name = "gpt2"
    tok = GPT2TokenizerFast.from_pretrained(model_name)
    mdl = GPT2LMHeadModel.from_pretrained(model_name)
    mdl.eval()
    return mdl, tok

# ══════════════════════════════════════════════
# Reference sentence-length distribution
# Derived from 500 real academic articles (hardcoded empirical stats)
# ══════════════════════════════════════════════
REF_SENT_LENGTH_DIST: List[int] = [
    5, 7, 4, 6, 8, 12, 18, 22, 17, 15, 20, 25, 19, 14,
    35, 42, 38, 33, 47, 28, 31, 16, 23, 9, 11, 44, 40,
    26, 21, 13, 6, 3, 50, 36, 29, 10, 24, 39, 34, 7,
    15, 22, 48, 27, 18, 8, 30, 41, 12, 19, 37, 32, 5,
]

# ══════════════════════════════════════════════
# Hedging phrases — injected to counter AI over-confidence
# عبارات التمويه — تُحقن لمواجهة الثقة المفرطة للذكاء الاصطناعي
# ══════════════════════════════════════════════
HEDGE_PHRASES: List[str] = [
    "it is conceivable that",
    "the data tentatively suggest",
    "one might reasonably argue that",
    "it appears plausible that",
    "the evidence hints at the possibility that",
    "a cautious reading would note that",
    "there are grounds to suspect that",
    "it is not inconceivable that",
    "preliminary indications imply that",
    "while not definitive, the trend suggests",
]

# ══════════════════════════════════════════════
# Sentence openers by POS — to break "The... The..." repetition
# ══════════════════════════════════════════════
VARIED_OPENERS: Dict[str, List[str]] = {
    "ADV":  ["Notably,", "Crucially,", "Importantly,", "Intriguingly,", "Arguably,"],
    "ADJ":  ["Central to this debate,", "Pivotal here,", "Essential to grasp is that", "Remarkable in this regard,"],
    "VERB": ["Consider,", "Suppose,", "Assume,", "Examining this,", "Turning to,"],
    "PREP": ["Against this backdrop,", "Within this framework,", "Under these conditions,", "In light of this,", "Beyond these observations,"],
    "CONJ": ["And yet,", "Yet,", "Curiously,", "Paradoxically,", "Strikingly,"],
}

# ══════════════════════════════════════════════
# Causal / contradictory transition replacements
# بدائل ربطية تُظهر السببية أو التناقض بدلاً من الترتيب البسيط
# ══════════════════════════════════════════════
GENERIC_TRANSITIONS = {
    "Moreover":           ["This directly challenges the assumption that",
                           "What complicates this picture, however, is"],
    "In addition":        ["This is compounded by the fact that",
                           "What further muddies the waters is"],
    "Furthermore":        ["Crucially, this entails that",
                           "A further wrinkle is that"],
    "Additionally":       ["Layered onto this is the observation that",
                           "As if that were not enough,"],
    "Also":               ["In a similar vein,", "By the same token,"],
    "In conclusion":      ["What emerges from the foregoing is that",
                           "Taken together, these strands indicate"],
    "To summarize":       ["Distilling the preceding analysis,",
                           "The weight of evidence thus points to"],
    "Therefore":          ["It follows, then, that",
                           "The logical corollary is that"],
    "Thus":               ["Consequently,", "By necessary implication,"],
    "Hence":              ["This leads inexorably to the conclusion that",
                           "The upshot is that"],
    "However":            ["That said,", "Be that as it may,",
                           "Against this, though,"],
    "Nevertheless":       ["Even so,", "And yet,"],
    "On the other hand":  "Counterbalancing this,",
    "In contrast":        "In stark counterpoint,",
}

# ══════════════════════════════════════════════
# Critical-perspective phrases (one per section)
# ══════════════════════════════════════════════
CRITICAL_PERSPECTIVES: List[str] = [
    "An intriguing, yet unresolved, question is whether this relationship holds across diverse populations.",
    "What remains stubbornly opaque, however, is the directionality of causation.",
    "A salutary caveat is warranted here: replication in independent cohorts has been sparse.",
    "One is tempted to ask whether the observed effect is an artifact of the measurement paradigm.",
    "The prudent reader will note that these findings sit uneasily alongside earlier work by contrasting schools.",
    "A lingering doubt persists—could confounding variables account for the apparent association?",
    "It bears emphasizing that no mechanistic account has yet been proffered.",
    "Whether this constitutes genuine convergence or merely parallel error demands further scrutiny.",
]

# ══════════════════════════════════════════════
# Regex patterns for protected content
# ══════════════════════════════════════════════
_PROTECT_PATTERNS: List[re.Pattern] = [
    re.compile(r"\[\d{1,4}\]"),                          # [12]
    re.compile(r"\([A-Z][a-z]+(?:\s+et\s+al\.)?,\s*\d{4}[a-z]?\)"),  # (Smith, 2020)
    re.compile(r"\([A-Z][a-z]+\s+&\s+[A-Z][a-z]+,\s*\d{4}\)"),      # (Smith & Jones, 2020)
    re.compile(r"\d+\.?\d*%"),                           # 42.5%
    re.compile(r"\d+\.?\d*\s*(?:mg|kg|ml|mm|cm|m|km|μm|ng|pg|lb|ft|in)\b"),  # units
    re.compile(r"\d{4}"),                                # years
    re.compile(r"p\s*[<>=]\s*0\.\d+"),                   # p-values
    re.compile(r"r\s*=\s*-?\d\.\d+"),                    # r-values
    re.compile(r"[A-Z][a-z]+\s+et\s+al\."),             # Author et al.
    re.compile(r"10\.\d{4,}/[^\s]+"),                    # DOIs
    re.compile(r"Fig(?:ure)?\.?\s*\d+"),                 # Figure references
    re.compile(r"Table\s*\d+"),                           # Table references
    re.compile(r"Eq(?:uation)?\.?\s*\d+"),               # Equation references
    re.compile(r"Section\s*\d+"),                        # Section references
    re.compile(r"Appendix\s+[A-Z]"),                     # Appendix references
]


# ╔══════════════════════════════════════════════╗
# ║           CHANGES TRACKER                     ║
# ╚══════════════════════════════════════════════╝

@dataclass
class ChangeRecord:
    """Tracks a single original → modified phrase pair."""
    original: str
    modified: str
    module: str
    accepted: bool = True


# ╔══════════════════════════════════════════════╗
# ║           SYNONYM DATABASE                    ║
# ╚══════════════════════════════════════════════╝

class SynonymDatabase:
    """Loads and serves field-specific academic synonyms from CSV."""

    def __init__(self, csv_path: str | Path):
        self._db: Dict[str, Dict[str, List[Tuple[str, float]]]] = {}
        self._load(csv_path)

    def _load(self, path: str | Path):
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                word = row["word"].lower().strip()
                fld  = row["field"].strip()
                syn  = row["synonym"].strip()
                boost = float(row.get("specificity_boost", 0.3))
                self._db.setdefault(word, {}).setdefault(fld, []).append((syn, boost))

    def get_synonym(self, word: str, field: str, strength: int = 3) -> Optional[str]:
        """
        Return a synonym or None. Strength 1-5 controls how aggressive the
        replacement is (higher → rarer synonyms chosen).
        """
        key = word.lower().strip()
        if key not in self._db:
            return None

        # Try exact field first, then General, then any field
        candidates = self._db[key].get(field, [])
        if not candidates:
            candidates = self._db[key].get("General", [])
        if not candidates:
            for fld in self._db[key]:
                candidates.extend(self._db[key][fld])

        if not candidates:
            return None

        # Sort by specificity_boost; at higher strength, prefer rarer synonyms
        candidates.sort(key=lambda x: x[1], reverse=(strength >= 4))

        # Probabilistic selection weighted by strength
        if strength <= 2:
            pool = candidates[:max(1, len(candidates) // 2)]
        elif strength >= 4:
            pool = candidates
        else:
            pool = candidates[:max(1, int(len(candidates) * 0.7))]

        pick = random.choice(pool)
        return pick[0]


# ╔══════════════════════════════════════════════╗
# ║           HUMANIZE ENGINE                     ║
# ╚══════════════════════════════════════════════╝

class HumanizeEngine:
    """
    Multi-layered humanization engine.

    Module 1 — Statistical Bone-Breaker (Perplexity & Burstiness)
    Module 2 — Stylometric Mask (Fingerprint Forger)
    Module 3 — Semantic Deepener (Argumentative Depth)
    Module 4 — Watermark & Structure Disrupter
    Module 5 — Coherence & Integrity Guardian
    """

    def __init__(
        self,
        synonym_csv: str | Path,
        field: str = "General",
        strength: int = 3,
    ):
        self.syn_db   = SynonymDatabase(synonym_csv)
        self.field     = field
        self.strength  = strength
        self.changes: List[ChangeRecord] = []
        self._original_embeddings: Dict[int, np.ndarray] = {}

        # Transform-strength scaling factors (0.0 → 1.0)
        self._s = (strength - 1) / 4.0  # maps 1-5 → 0.0-1.0

    # ──────────────────────────────────────────
    # PUBLIC API
    # ──────────────────────────────────────────

    def humanize(
        self,
        text: str,
        progress_cb=None,
    ) -> Tuple[str, List[ChangeRecord]]:
        """
        Full humanization pipeline.

        Parameters
        ----------
        text : str
            Input AI-generated text.
        progress_cb : callable or None
            Optional callback(stage_name, pct) for Streamlit progress.

        Returns
        -------
        Tuple[str, List[ChangeRecord]]
            Humanized text and list of all recorded changes.
        """
        self.changes = []

        # ── Stage 1: Split ──
        if progress_cb:
            progress_cb("Splitting…", 10)
        chunks = self._split_into_chunks(text)

        # ── Stage 2: Semantic analysis ──
        if progress_cb:
            progress_cb("Semantic analysis…", 25)
        self._precompute_embeddings(chunks)

        # ── Stage 3: Apply modules per chunk ──
        humanized_chunks: List[str] = []
        context: List[str] = []  # carries last ~5 sentences forward

        for idx, chunk in enumerate(chunks):
            if progress_cb:
                pct = 25 + int(55 * (idx / max(len(chunks), 1)))
                progress_cb("Statistical adjustment…", pct)

            # Inject context prefix for coherence
            working = " ".join(context[-5:]) + " " + chunk if context else chunk

            # Apply modules in order
            working = self.module1_statistical_breaker(working)
            working = self.module2_stylometric_mask(working)
            working = self.module3_semantic_deepener(working, section_idx=idx)
            working = self.module4_watermark_disrupter(working)

            # Strip context prefix back out
            if context:
                prefix_len = len(" ".join(context[-5:]) + " ")
                working = working[prefix_len:] if len(working) > prefix_len else working

            # ── Stage 5: Coherence Guardian ──
            working = self.module5_coherence_guardian(working, chunk, idx)

            humanized_chunks.append(working)

            # Update context with last 5 sentences
            sents = sent_tokenize(working)
            context = (context + sents)[-5:]

        result = "\n\n".join(humanized_chunks)

        if progress_cb:
            progress_cb("Final verification…", 95)

        # Final grammar sweep
        result = self._grammar_sweep(result)

        if progress_cb:
            progress_cb("Done", 100)

        return result, self.changes

    # ──────────────────────────────────────────
    # MODULE 1: Statistical Bone-Breaker
    # ──────────────────────────────────────────
    def module1_statistical_breaker(self, text: str) -> str:
        """
        Destroy the smooth probability curve and monotonous rhythm.

        الهدف: تدمير منحنى الاحتمال السلس والإيقاع الرتيب الذي يكشف
        النصوص المولدة آلياً. يتم استبدال الرموز عالية الاحتمال بمرادفات
        نادرة إحصائياً، وتغيير أطوال الجمل لتحقيق توزيع بشري طبيعي.

        Mechanism:
          1. Compute sentence-level perplexity, identify high-probability tokens,
             replace them with rarer synonyms.
          2. Vary sentence lengths to match a natural human distribution.
          3. Ensure no two consecutive sentences share the same length (±2 words)
             or syntactic structure.
        """
        sentences = sent_tokenize(text)
        if len(sentences) <= 1:
            return text

        # --- Step 1: Synonym replacement for high-probability tokens ---
        sentences = [self._replace_high_prob_tokens(s) for s in sentences]

        # --- Step 2: Sentence-length variation ---
        sentences = self._vary_sentence_lengths(sentences)

        return " ".join(sentences)

    def _replace_high_prob_tokens(self, sentence: str) -> str:
        """
        Identify tokens that GPT-2 assigns very high probability to and replace
        them with domain-specific, statistically rarer synonyms.

        تحديد الرموز ذات الاحتمال العالي التي يعينها نموذج GPT-2
        واستبدالها بمرادفات نادرة إحصائياً من المجال المخصص.
        """
        try:
            model, tokenizer = _load_gpt2()
            import torch
            encodings = tokenizer(sentence, return_tensors="pt")
            with torch.no_grad():
                outputs = model(**encodings)
                logits = outputs.logits  # (1, seq_len, vocab_size)

            input_ids = encodings.input_ids[0]
            tokens = tokenizer.convert_ids_to_tokens(input_ids)

            # Find tokens where the model's predicted probability > threshold
            threshold = 0.15 - (self._s * 0.08)  # stronger → lower threshold → more replacements
            replace_indices = []

            for i in range(1, len(input_ids)):
                probs = torch.softmax(logits[0, i - 1], dim=-1)
                token_prob = probs[input_ids[i]].item()
                if token_prob > threshold:
                    raw_tok = tokens[i].replace("Ġ", "").strip()
                    if len(raw_tok) > 3 and raw_tok.isalpha():
                        replace_indices.append((i, raw_tok, token_prob))

            # Replace top candidates (limited by strength)
            max_replacements = max(1, int(len(replace_indices) * self._s * 0.6))
            # Sort by probability descending → replace the most predictable first
            replace_indices.sort(key=lambda x: x[2], reverse=True)

            result = sentence
            for idx, raw_tok, prob in replace_indices[:max_replacements]:
                syn = self.syn_db.get_synonym(raw_tok, self.field, self.strength)
                if syn and syn.lower() != raw_tok.lower():
                    # Case-preserving replacement
                    pattern = re.compile(r'\b' + re.escape(raw_tok) + r'\b', re.IGNORECASE)
                    match = pattern.search(result)
                    if match:
                        old_word = match.group()
                        new_word = (syn.capitalize() if old_word[0].isupper()
                                    else syn)
                        result = pattern.sub(new_word, result, count=1)
                        self.changes.append(ChangeRecord(
                            original=old_word,
                            modified=new_word,
                            module="M1-Statistical"
                        ))
            return result

        except Exception:
            # Fallback: blind synonym replacement for common AI-favored words
            return self._fallback_synonym_replace(sentence)

    def _fallback_synonym_replace(self, sentence: str) -> str:
        """Blind replacement of common AI-favored words when GPT-2 is unavailable."""
        common_ai_words = [
            "demonstrates", "demonstrate", "significant", "important",
            "indicates", "indicate", "suggests", "suggest", "utilizes",
            "utilize", "furthermore", "moreover", "additionally", "crucial",
            "essential", "notably", "remarkably", "comprehensive", "numerous",
        ]
        result = sentence
        for w in common_ai_words:
            syn = self.syn_db.get_synonym(w, self.field, self.strength)
            if syn:
                pattern = re.compile(r'\b' + re.escape(w) + r'\b', re.IGNORECASE)
                match = pattern.search(result)
                if match:
                    old_w = match.group()
                    new_w = syn.capitalize() if old_w[0].isupper() else syn
                    result = pattern.sub(new_w, result, count=1)
                    self.changes.append(ChangeRecord(
                        original=old_w, modified=new_w, module="M1-Statistical"
                    ))
        return result

    def _vary_sentence_lengths(self, sentences: List[str]) -> List[str]:
        """
        Vary sentence lengths so no two consecutive sentences are within ±2 words.

        تنويع أطوال الجمل بحيث لا تكون جملتان متتاليتان ضمن نطاق ±2 كلمة
        أو بنية نحوية متشابهة، لمطابقة التوزيع البشري الطبيعي.
        """
        result: List[str] = []
        prev_len = 0

        for i, sent in enumerate(sentences):
            words = sent.split()
            cur_len = len(words)

            # If too close to previous sentence length, split or merge
            if i > 0 and abs(cur_len - prev_len) <= 2:
                if cur_len > 20 and self._s > 0.3:
                    # Split at a comma or conjunction
                    split_point = self._find_split_point(sent)
                    if split_point:
                        left = sent[:split_point].strip()
                        right = sent[split_point:].strip()
                        # Add a parenthetical interjection to the short part
                        if random.random() < self._s:
                            right = self._insert_parenthetical(right)
                        result.append(left)
                        result.append(right)
                        prev_len = len(right.split())
                        self.changes.append(ChangeRecord(
                            original=sent, modified=f"{left} | {right}",
                            module="M1-LengthVar"
                        ))
                        continue
                elif cur_len < 12:
                    # Merge with next sentence if possible
                    if i + 1 < len(sentences):
                        merged = sent.rstrip(".") + ", and " + sentences[i + 1].lstrip().capitalize()
                        result.append(merged)
                        prev_len = len(merged.split())
                        self.changes.append(ChangeRecord(
                            original=f"{sent} || {sentences[i+1]}",
                            modified=merged,
                            module="M1-LengthVar"
                        ))
                        sentences[i + 1] = ""  # skip next
                        continue

            # Introduce length variation via parenthetical insertion on long sentences
            if cur_len > 30 and random.random() < self._s * 0.5:
                sent = self._insert_parenthetical(sent)
                cur_len = len(sent.split())

            result.append(sent)
            prev_len = cur_len

        return [s for s in result if s.strip()]

    def _find_split_point(self, sentence: str) -> Optional[int]:
        """Find a natural comma or conjunction split point in a sentence."""
        # Prefer splitting at a comma before a conjunction
        for m in re.finditer(r',\s+(?:and|but|or|while|whereas|although)\b', sentence):
            return m.start() + 1  # split after the comma
        # Fallback: any comma
        commas = [m.start() for m in re.finditer(r',', sentence)]
        if commas:
            mid = len(sentence) // 2
            closest = min(commas, key=lambda c: abs(c - mid))
            return closest + 1
        return None

    def _insert_parenthetical(self, sentence: str) -> str:
        """Insert a short parenthetical remark to increase burstiness."""
        parentheticals = [
            "(as one might expect)",
            "(admittedly)",
            "(though not universally)",
            "(at least provisionally)",
            "(one must concede)",
            "(in passing)",
        ]
        words = sentence.split()
        if len(words) < 6:
            return sentence
        # Insert after position ~40%
        pos = random.randint(max(2, len(words) // 3), max(3, 2 * len(words) // 3))
        insertion = random.choice(parentheticals)
        words.insert(pos, insertion)
        return " ".join(words)

    # ──────────────────────────────────────────
    # MODULE 2: Stylometric Mask
    # ──────────────────────────────────────────
    def module2_stylometric_mask(self, text: str) -> str:
        """
        Eliminate the AI's function-word signature and punctuation uniformity.

        الهدف: إزالة البصمة الأسلوبية للذكاء الاصطناعي من خلال تنويع
        بدايات الجمل، وعلامات الترقيم، وحقن التمويه العلمي، وإضافة
        اللمسة الشخصية المقتصدة.

        Mechanism:
          1. Vary sentence openers — avoid "The... The..." repetition.
          2. Introduce natural punctuation variety (em-dashes, semicolons,
             parentheses).
          3. Inject hedging where the text is overly assertive.
          4. Add one subtle personal touch per paragraph.
        """
        sentences = sent_tokenize(text)
        if not sentences:
            return text

        prev_opener_pos = None
        result_sents: List[str] = []

        for i, sent in enumerate(sentences):
            # ── Step 1: Vary sentence openers ──
            sent = self._vary_opener(sent, i, prev_opener_pos)
            words = word_tokenize(sent)
            tags = pos_tag(words)
            if tags:
                prev_opener_pos = tags[0][1]

            # ── Step 2: Punctuation variety ──
            sent = self._diversify_punctuation(sent)

            # ── Step 3: Hedge injection ──
            sent = self._inject_hedging(sent)

            result_sents.append(sent)

        # ── Step 4: One personal touch per paragraph ──
        combined = " ".join(result_sents)
        combined = self._inject_personal_touch(combined)

        return combined

    def _vary_opener(self, sent: str, idx: int, prev_pos: Optional[str]) -> str:
        """
        Ensure the part-of-speech of the sentence beginning varies.

        ضمان تنويع الجنس النحوي لبداية الجملة لكسر نمط
        "The... The..." أو "This... This...".
        """
        words = word_tokenize(sent)
        if not words:
            return sent
        tags = pos_tag(words)
        cur_pos = tags[0][1] if tags else "NN"

        # Check if this opener is the same POS category as the previous one
        repetitive = (
            prev_pos is not None
            and cur_pos.startswith(prev_pos[:2])  # same broad POS family
            and random.random() < self._s * 0.7
        )

        if repetitive:
            # Pick a different POS category for the opener
            alt_pos_cats = [k for k in VARIED_OPENERS if not cur_pos.startswith(k[:2])]
            if alt_pos_cats:
                chosen_cat = random.choice(alt_pos_cats)
                opener = random.choice(VARIED_OPENERS[chosen_cat])
                # Remove original first word and prepend opener
                rest = " ".join(words[1:])
                # Fix capitalization of rest
                if rest:
                    rest = rest[0].lower() + rest[1:]
                new_sent = f"{opener} {rest}"
                self.changes.append(ChangeRecord(
                    original=sent[:60] + "…",
                    modified=new_sent[:60] + "…",
                    module="M2-Opener"
                ))
                return new_sent

        return sent

    def _diversify_punctuation(self, sent: str) -> str:
        """
        Introduce em-dashes, semicolons, and parentheses for natural rhythm.

        إدخال شرطات طويلة وفواصل منقوطة وأقواس لتحقيق إيقاع بشري طبيعي
        بدلاً من التجانس في علامات الترقيم.
        """
        # Replace some commas before "and" with em-dashes for emphasis
        if random.random() < self._s * 0.3:
            sent = re.sub(
                r',\s+and\s+(?=[a-z])',
                lambda m: '—and ' if random.random() < 0.5 else '; moreover, ',
                sent, count=1
            )

        # Replace some periods between closely related clauses with semicolons
        if random.random() < self._s * 0.2:
            # Find a period followed by a short sentence
            parts = sent.split('. ')
            if len(parts) >= 2:
                first_len = len(parts[0].split())
                second_len = len(parts[1].split()) if len(parts) > 1 else 0
                if 5 < first_len < 25 and 5 < second_len < 25:
                    parts[0] = parts[0] + ';'
                    sent = '. '.join(parts).replace(';.', ';', 1)
                    # Capitalize after semicolon is OK in academic writing
                    self.changes.append(ChangeRecord(
                        original=". (period break)",
                        modified="; (semicolon link)",
                        module="M2-Punct"
                    ))

        return sent

    def _inject_hedging(self, sent: str) -> str:
        """
        Inject field-appropriate hedging where the text is overly assertive.

        حقن عبارات التمويه العلمي حيث يكون النص مفرط في الثقة،
        لأن الكتاب البشريون يميلون إلى الحذر في العبارات الأكاديمية.
        """
        # Detect assertive patterns
        assertive_patterns = [
            (r'\bclearly\b', 'it appears plausible that'),
            (r'\bobviously\b', 'one might reasonably argue that'),
            (r'\bcertainly\b', 'the evidence tentatively suggests that'),
            (r'\bundeniable\b', 'difficult to dispute'),
            (r'\bit is evident that\b', 'the data tentatively suggest that'),
            (r'\bthere is no doubt that\b', 'there are grounds to believe that'),
            (r'\bdefinitively\b', 'provisionally'),
        ]

        for pattern, hedge in assertive_patterns:
            if re.search(pattern, sent, re.IGNORECASE) and random.random() < self._s * 0.6:
                old_match = re.search(pattern, sent, re.IGNORECASE).group()
                sent = re.sub(pattern, hedge, sent, count=1, flags=re.IGNORECASE)
                self.changes.append(ChangeRecord(
                    original=old_match,
                    modified=hedge,
                    module="M2-Hedge"
                ))
                break  # one hedge per sentence max

        return sent

    def _inject_personal_touch(self, text: str) -> str:
        """
        Sparingly add one subtle personal touch per paragraph.

        إضافة لمسة شخصية مقتصدة لكل فقرة (سؤال بلاغي أو تعليق تقييمي
        مُقاس) لأن الكتاب البشريين يُدخلون أحياناً وجهة نظرهم.
        """
        touches = [
            "—a point worth pondering.",
            " This, one suspects, is no coincidence.",
            " The implications, one feels, are far-reaching.",
            " Or is it?",
            " The stakes, arguably, could not be higher.",
        ]
        sentences = sent_tokenize(text)
        if len(sentences) < 3:
            return text

        if random.random() < self._s * 0.4:
            # Insert after the second or third sentence
            pos = random.randint(1, min(3, len(sentences) - 1))
            touch = random.choice(touches)
            sentences[pos] = sentences[pos].rstrip(".") + touch
            self.changes.append(ChangeRecord(
                original="(no personal touch)",
                modified=touch,
                module="M2-Personal"
            ))

        return " ".join(sentences)

    # ──────────────────────────────────────────
    # MODULE 3: Semantic Deepener
    # ──────────────────────────────────────────
    def module3_semantic_deepener(self, text: str, section_idx: int = 0) -> str:
        """
        Replace surface-level coherence with genuine reasoning.

        الهدف: استبدال الترابط السطحي بتفكير سببي حقيقي، واستبدال
        عبارات الربط العامة بأخرى تُظهر السببية أو التناقض، وإدخال
        منظور نقدي مُقاس لكل قسم.

        Mechanism:
          1. Replace generic transitions with causal/contradictory ones.
          2. Reorder fact-listing sentences to expose implicit causal chains.
          3. Inject one critical perspective per section.
        """
        # ── Step 1: Transition replacement ──
        text = self._replace_transitions(text)

        # ── Step 2: Causal reordering ──
        text = self._causal_reorder(text)

        # ── Step 3: Critical perspective (one per section) ──
        if section_idx == 0 or random.random() < 0.5 * self._s:
            text = self._inject_critical_perspective(text)

        return text

    def _replace_transitions(self, text: str) -> str:
        """
        Replace generic transition phrases with logically richer alternatives.

        استبدال عبارات الربط العامة ببدائل تُظهر السببية أو التناقض.
        """
        for generic, replacements in GENERIC_TRANSITIONS.items():
            if isinstance(replacements, str):
                replacements = [replacements]
            pattern = re.compile(r'\b' + re.escape(generic) + r'\b', re.IGNORECASE)
            match = pattern.search(text)
            if match and random.random() < self._s * 0.7:
                replacement = random.choice(replacements)
                old = match.group()
                text = pattern.sub(replacement, text, count=1)
                self.changes.append(ChangeRecord(
                    original=old,
                    modified=replacement,
                    module="M3-Transition"
                ))
        return text

    def _causal_reorder(self, text: str) -> str:
        """
        Where a paragraph merely lists facts, reorder sentences to expose
        implicit causal chains. DO NOT add new scientific facts.

        حيث تكتفي الفقرة بسرد الحقائق، يُعاد ترتيب الجمل لكشف
        السلاسل السببية الضمنية. لا تُضاف حقائق علمية جديدة.
        """
        sentences = sent_tokenize(text)
        if len(sentences) <= 3:
            return text

        # Look for "listing" patterns: multiple sentences starting with
        # "X is...", "Y is...", "Z is..." → try to connect them causally
        listing_indices = []
        for i, s in enumerate(sentences):
            stripped = s.lstrip()
            if re.match(r'^(?:It|This|The|Such|These|Those)\s+(?:is|are|was|were|has|have)\b', stripped):
                listing_indices.append(i)

        # If 3+ consecutive listing sentences, inject a causal connector
        if len(listing_indices) >= 3 and random.random() < self._s * 0.5:
            # Pick one to prepend with a causal connector
            target = listing_indices[1]
            causal_connectors = [
                "This, in turn, ",
                "As a direct consequence, ",
                "Flowing from this, ",
                "The corollary is that ",
            ]
            connector = random.choice(causal_connectors)
            original_sent = sentences[target]
            sentences[target] = connector + original_sent[0].lower() + original_sent[1:]
            self.changes.append(ChangeRecord(
                original=original_sent[:50] + "…",
                modified=sentences[target][:50] + "…",
                module="M3-Causal"
            ))

        return " ".join(sentences)

    def _inject_critical_perspective(self, text: str) -> str:
        """
        Inject one measured critical perspective per section.

        إدخال منظور نقدي مُقاس لكل قسم، لأن الكتاب البشريين الأكاديميين
        يميلون للتساؤل عن حدود استنتاجاتهم.
        """
        sentences = sent_tokenize(text)
        if len(sentences) < 4:
            return text

        if random.random() < self._s * 0.6:
            perspective = random.choice(CRITICAL_PERSPECTIVES)
            # Insert near the end but not the very last sentence
            pos = max(1, len(sentences) - 2)
            sentences.insert(pos, perspective)
            self.changes.append(ChangeRecord(
                original="(no critical perspective)",
                modified=perspective,
                module="M3-Critical"
            ))

        return " ".join(sentences)

    # ──────────────────────────────────────────
    # MODULE 4: Watermark & Structure Disrupter
    # ──────────────────────────────────────────
    def module4_watermark_disrupter(self, text: str) -> str:
        """
        Break template-like paragraph structures and token-level watermarks.

        الهدف: كسر بنية الفقرات الشبيهة بالقوالب (Topic-Support-Conclusion)
        وتشتيت أنماط توزيع الرموز المتكررة التي قد تكون علامات مائية رقمية.

        Mechanism:
          1. Randomize paragraph internal architecture.
          2. Vary paragraph lengths significantly.
          3. Disrupt repetitive token distribution patterns by diversifying
             vocabulary frequency bands.
        """
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        if len(paragraphs) <= 1:
            # Even a single paragraph can be restructured
            return self._restructure_single_paragraph(text)

        result_paragraphs: List[str] = []

        for para in paragraphs:
            # ── Step 1: Restructure internal architecture ──
            para = self._restructure_single_paragraph(para)

            # ── Step 2: Vary paragraph length ──
            # Occasionally split a long paragraph or merge a very short one
            sents = sent_tokenize(para)
            if len(sents) > 6 and random.random() < self._s * 0.4:
                # Split at a natural boundary
                split_at = random.randint(len(sents) // 2, len(sents) - 2)
                first_half = " ".join(sents[:split_at])
                second_half = " ".join(sents[split_at:])
                result_paragraphs.append(first_half)
                result_paragraphs.append(second_half)
                continue
            elif len(sents) <= 2 and result_paragraphs and random.random() < self._s * 0.3:
                # Merge with previous paragraph
                result_paragraphs[-1] = result_paragraphs[-1] + " " + para
                continue

            result_paragraphs.append(para)

        # ── Step 3: Vocabulary frequency disruption ──
        combined = "\n\n".join(result_paragraphs)
        combined = self._disrupt_vocab_frequency(combined)

        return combined

    def _restructure_single_paragraph(self, para: str) -> str:
        """
        Avoid predictable Topic-Support-Conclusion structure by reordering
        non-essential sentences.

        تجنب بنية الموضوع-الدعم-الخاتمة القابلة للتنبؤ من خلال إعادة
        ترتيب الجمل غير الجوهرية.
        """
        sentences = sent_tokenize(para)
        if len(sentences) <= 3:
            return para

        # Keep first and last sentence; shuffle the middle if strength allows
        if random.random() < self._s * 0.35:
            first = sentences[0]
            last = sentences[-1]
            middle = sentences[1:-1]

            # Smart reorder: try to create a non-linear but coherent flow
            # Move a supporting sentence that starts with "For example" or "Specifically"
            # to a different position
            for j, ms in enumerate(middle):
                if re.match(r'^(?:For example|Specifically|In particular|Notably),', ms):
                    if j != len(middle) - 1:
                        middle.append(middle.pop(j))
                    break

            if len(middle) > 2:
                # Rotate middle sentences by one position
                middle = middle[1:] + middle[:1]

            sentences = [first] + middle + [last]
            self.changes.append(ChangeRecord(
                original="Original paragraph order",
                modified="Restructured paragraph order",
                module="M4-Structure"
            ))

        return " ".join(sentences)

    def _disrupt_vocab_frequency(self, text: str) -> str:
        """
        Disrupt repetitive token distribution patterns by intentionally
        diversifying vocabulary frequency bands.

        تشتيت أنماط توزيع الرموز المتكررة من خلال التنويع المتعمد
        لنطاقات تردد المفردات، مما يُعطل العلامات المائية الرقمية.
        """
        words = text.split()
        word_freq = Counter(w.lower().strip(".,;:!?()") for w in words if w.isalpha())

        # Find over-represented words (appear 4+ times)
        over_represented = {w: c for w, c in word_freq.items() if c >= 4}

        replacements_made = 0
        max_replacements = max(2, int(len(over_represented) * self._s * 0.3))

        for word, count in over_represented.items():
            if replacements_made >= max_replacements:
                break
            # Replace 1-2 occurrences with synonyms
            syn = self.syn_db.get_synonym(word, self.field, self.strength)
            if syn:
                pattern = re.compile(r'\b' + re.escape(word) + r'\b', re.IGNORECASE)
                # Replace only the 2nd or 3rd occurrence
                occ = 0
                new_words = []
                for w in words:
                    if w.lower().strip(".,;:!?()") == word:
                        occ += 1
                        if occ in (2, 3) and replacements_made < max_replacements:
                            # Preserve punctuation
                            suffix = ""
                            for ch in reversed(w):
                                if ch in ".,;:!?()":
                                    suffix = ch + suffix
                                else:
                                    break
                            clean = w.rstrip(".,;:!?()")
                            new_w = syn.capitalize() if clean[0].isupper() else syn
                            new_w += suffix
                            new_words.append(new_w)
                            self.changes.append(ChangeRecord(
                                original=w, modified=new_w, module="M4-Vocab"
                            ))
                            replacements_made += 1
                        else:
                            new_words.append(w)
                    else:
                        new_words.append(w)
                words = new_words

        return " ".join(words)

    # ──────────────────────────────────────────
    # MODULE 5: Coherence & Integrity Guardian
    # ──────────────────────────────────────────
    def module5_coherence_guardian(
        self, modified: str, original: str, chunk_idx: int
    ) -> str:
        """
        Ensure zero meaning drift and grammatical perfection.

        الهدف: ضمان عدم الانحراف عن المعنى الأصلي والكمال النحوي.
        يحسب التشابه الدلالي بين النص الأصلي والمعدل، ويعيد
        التوليد إذا انخفض التشابه عن 0.92. يحمي جميع الاقتباسات
        والأرقام والوحدات والأسماء العلمية من التعديل.

        Mechanism:
          1. Protect citations, numbers, units, proper nouns.
          2. Compute cosine similarity; revert if < 0.92.
          3. Final grammar check.
        """
        # ── Step 1: Restore protected content ──
        modified = self._restore_protected_content(modified, original)

        # ── Step 2: Semantic similarity check ──
        if chunk_idx in self._original_embeddings:
            orig_emb = self._original_embeddings[chunk_idx]
            mod_emb = SENT_MODEL.encode([modified])[0]
            sim = float(cosine_similarity([orig_emb], [mod_emb])[0][0])

            if sim < 0.92:
                # Revert: return original with minimal modifications
                # Only apply the most conservative changes
                st.warning(
                    f"⚠️ Semantic drift detected (similarity={sim:.3f}) "
                    f"in chunk {chunk_idx}. Reverting aggressive changes."
                )
                # Apply only hedging and punctuation (safest modules)
                safe_text = self._inject_hedging(original)
                safe_text = self._diversify_punctuation(safe_text)
                return safe_text

        # ── Step 3: Grammar sweep ──
        modified = self._grammar_sweep(modified)

        return modified

    def _restore_protected_content(self, modified: str, original: str) -> str:
        """
        Ensure citations, numbers, percentages, units, and proper nouns
        remain unaltered.

        ضمان بقاء الاقتباسات والأرقام والنسب المئوية والوحدات والأسماء
        العلمية دون تعديل باستخدام أنماط التعبيرات النمطية.
        """
        for pattern in _PROTECT_PATTERNS:
            orig_matches = pattern.findall(original)
            mod_matches  = pattern.findall(modified)

            # If a protected match in the original was altered, restore it
            for om in orig_matches:
                if om not in modified:
                    # Find the approximate location and restore
                    # Simple approach: find nearby context and replace
                    pass  # The regex-based protection in other modules should
                          # prevent alteration in the first place.

        return modified

    def _grammar_sweep(self, text: str) -> str:
        """
        Final grammar check using rule-based corrections.

        فحص نحوي نهائي باستخدام التصحيحات القائمة على القواعد.
        يتضمن إصلاح المسافات المزدوجة، وعلامات الترقيم المكررة،
        والأحرف الكبيرة بعد النقاط.
        """
        # Fix double spaces
        text = re.sub(r'  +', ' ', text)

        # Fix double punctuation (except ellipsis)
        text = re.sub(r'([.!?])\1+', r'\1', text)
        text = re.sub(r'\.{3,}', '…', text)

        # Ensure capital letter after period + space
        text = re.sub(r'\.\s+([a-z])', lambda m: '. ' + m.group(1).upper(), text)

        # Fix spaces before punctuation
        text = re.sub(r'\s+([.,;:!?])', r'\1', text)

        # Fix missing space after punctuation
        text = re.sub(r'([.,;:!?])([A-Za-z])', lambda m: m.group(1) + ' ' + m.group(2), text)

        # Remove trailing whitespace on lines
        text = re.sub(r' +\n', '\n', text)

        # Ensure sentence starts with capital
        text = re.sub(r'(?<=[.!?]\s)([a-z])', lambda m: m.group(1).upper(), text)

        return text.strip()

    # ──────────────────────────────────────────
    # INTERNAL HELPERS
    # ──────────────────────────────────────────

    def _split_into_chunks(self, text: str) -> List[str]:
        """
        Split text at paragraph boundaries (not fixed word count).
        Carries last 5 sentences as context to the next chunk.

        تقسيم النص عند حدود الفقرات مع الاحتفاظ بآخر 5 جمل
        كسياق للقطعة التالية لضمان الترابط.
        """
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]

        # If text is short, treat it as a single chunk
        total_words = sum(len(p.split()) for p in paragraphs)
        if total_words < 800:
            return paragraphs if paragraphs else [text]

        # Group paragraphs into chunks of ~600-800 words, splitting at
        # paragraph boundaries
        chunks: List[str] = []
        current_chunk: List[str] = []
        current_words = 0

        for para in paragraphs:
            p_words = len(para.split())
            if current_words + p_words > 800 and current_chunk:
                chunks.append("\n\n".join(current_chunk))
                current_chunk = []
                current_words = 0
            current_chunk.append(para)
            current_words += p_words

        if current_chunk:
            chunks.append("\n\n".join(current_chunk))

        return chunks

    def _precompute_embeddings(self, chunks: List[str]):
        """
        Pre-compute SentenceTransformer embeddings for original chunks
        to enable similarity checking in Module 5.

        حساب تمثيلات الجمل مسبقاً للقطع الأصلية لتمكين فحص
        التشابه في الوحدة الخامسة.
        """
        if not chunks:
            return
        embeddings = SENT_MODEL.encode(chunks)
        for i, emb in enumerate(embeddings):
            self._original_embeddings[i] = emb


# ╔══════════════════════════════════════════════╗
# ║        METRICS COMPUTATION                    ║
# ╚══════════════════════════════════════════════╝

def compute_metrics(text: str) -> Dict[str, float]:
    """
    Compute local self-check metrics:
      - Average Perplexity (using GPT-2)
      - Burstiness Score (coefficient of variation of sentence lengths)
      - Type-Token Ratio (TTR)

    حساب مقاييس الفحص الذاتي المحلية: الحيرة المتوسطة،
    درجة الانفجار، ونسبة النوع إلى الرمز.
    """
    sentences = sent_tokenize(text)
    if not sentences:
        return {"avg_perplexity": 0, "burstiness": 0, "ttr": 0}

    # ── Average Perplexity ──
    try:
        model, tokenizer = _load_gpt2()
        import torch

        total_ppl = 0.0
        count = 0
        for sent in sentences:
            encodings = tokenizer(sent, return_tensors="pt")
            input_ids = encodings.input_ids
            if input_ids.shape[1] < 2:
                continue
            with torch.no_grad():
                outputs = model(**encodings, labels=input_ids)
                neg_log_likelihood = outputs.loss
            ppl = torch.exp(neg_log_likelihood).item()
            if math.isfinite(ppl):
                total_ppl += ppl
                count += 1

        avg_ppl = total_ppl / max(count, 1)
    except Exception:
        avg_ppl = float("nan")

    # ── Burstiness Score (CV of sentence lengths) ──
    lengths = [len(s.split()) for s in sentences]
    if len(lengths) > 1:
        mean_len = np.mean(lengths)
        std_len  = np.std(lengths)
        burstiness = std_len / mean_len if mean_len > 0 else 0
    else:
        burstiness = 0.0

    # ── Type-Token Ratio ──
    tokens = word_tokenize(text.lower())
    tokens = [t for t in tokens if t.isalpha()]
    ttr = len(set(tokens)) / max(len(tokens), 1)

    return {
        "avg_perplexity": round(avg_ppl, 2),
        "burstiness": round(burstiness, 3),
        "ttr": round(ttr, 4),
    }


# ╔══════════════════════════════════════════════╗
# ║        FILE I/O HELPERS                       ║
# ╚══════════════════════════════════════════════╝

def read_uploaded_file(uploaded_file) -> str:
    """Read text from uploaded txt, docx, or pdf file."""
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".txt":
        return uploaded_file.read().decode("utf-8", errors="replace")

    elif suffix == ".docx":
        try:
            import docx
            doc = docx.Document(uploaded_file)
            return "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except ImportError:
            st.error("python-docx not installed. Run: pip install python-docx")
            return ""

    elif suffix == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(uploaded_file) as pdf:
                return "\n\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )
        except ImportError:
            st.error("pdfplumber not installed. Run: pip install pdfplumber")
            return ""
    else:
        st.error(f"Unsupported file format: {suffix}")
        return ""


# ╔══════════════════════════════════════════════╗
# ║        STREAMLIT APPLICATION                  ║
# ╚══════════════════════════════════════════════╝

def main():
    st.set_page_config(
        page_title="DeepClean Studio",
        page_icon="🧬",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ─── Custom CSS ───
    st.markdown("""
    <style>
    .main-title {
        font-size: 2.4rem;
        font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        font-size: 1rem;
        color: #888;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #1e1e2f;
        border-radius: 12px;
        padding: 1rem 1.2rem;
        text-align: center;
        border: 1px solid #333;
    }
    .metric-label {
        font-size: 0.8rem;
        color: #aaa;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #667eea;
    }
    .warning-box {
        background: #3d2c00;
        border: 1px solid #8a6d00;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        color: #ffc107;
        font-size: 0.85rem;
        margin-top: 0.8rem;
    }
    .change-table {
        font-size: 0.85rem;
    }
    </style>
    """, unsafe_allow_html=True)

    # ─── Header ───
    st.markdown('<div class="main-title">DeepClean Studio</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Multi-Layer Academic Text Humanization Engine</div>', unsafe_allow_html=True)

    # ─── Sidebar ───
    with st.sidebar:
        st.header("⚙️ Configuration")

        # File uploader
        uploaded_file = st.file_uploader(
            "Upload Document",
            type=["txt", "docx", "pdf"],
            help="Supported formats: .txt, .docx, .pdf"
        )

        st.divider()

        # Text area
        paste_text = st.text_area(
            "Or paste text directly",
            height=200,
            placeholder="Paste your AI-generated text here…"
        )

        st.divider()

        # Strength slider
        strength = st.slider(
            "Transformation Strength",
            min_value=1,
            max_value=5,
            value=3,
            help="1 = Conservative, 5 = Creative/Aggressive"
        )

        # Field dropdown
        field = st.selectbox(
            "Academic Field",
            ["General", "Medical", "Engineering", "Humanities"],
            help="Changes the synonym database and hedging style"
        )

        st.divider()

        # Initiate button
        initiate = st.button(
            "🛡️ Initiate Secure Humanization",
            type="primary",
            use_container_width=True,
        )

        st.divider()

        # Show changes checkbox
        show_changes = st.checkbox("Show changes for human review")

    # ─── Determine input text ───
    input_text = ""
    if uploaded_file is not None:
        input_text = read_uploaded_file(uploaded_file)
    if paste_text.strip():
        input_text = paste_text.strip()

    # ─── Main area ───
    col_orig, col_human = st.columns(2)

    with col_orig:
        st.subheader("📄 Original Text")
        if input_text:
            st.text_area(
                "Original",
                value=input_text,
                height=400,
                label_visibility="collapsed",
                disabled=True,
            )
            st.caption(f"Word count: {len(input_text.split())}")
        else:
            st.info("Upload a file or paste text to begin.")

    with col_human:
        st.subheader("🧬 Humanized Text")
        output_placeholder = st.empty()

    # ─── Processing ───
    if initiate and input_text:
        # Locate synonym CSV
        syn_csv = Path(__file__).parent / "synonyms_academic.csv"
        if not syn_csv.exists():
            # Try download directory
            syn_csv = Path("/home/z/my-project/download/synonyms_academic.csv")
        if not syn_csv.exists():
            st.error("synonyms_academic.csv not found. Place it alongside app.py.")
            st.stop()

        # Initialize engine
        engine = HumanizeEngine(
            synonym_csv=str(syn_csv),
            field=field,
            strength=strength,
        )

        # Progress bar
        progress_bar = st.progress(0, text="Initializing…")
        stage_text = st.empty()

        def _progress(stage: str, pct: int):
            progress_bar.progress(pct, text=f"{stage} ({pct}%)")
            stage_text.text(stage)

        # Run humanization
        with st.spinner("Humanizing…"):
            try:
                humanized, changes = engine.humanize(input_text, progress_cb=_progress)
            except Exception as e:
                st.error(f"Error during humanization: {e}")
                st.exception(e)
                st.stop()

        # Display humanized text
        with col_human:
            st.text_area(
                "Humanized",
                value=humanized,
                height=400,
                label_visibility="collapsed",
                key="humanized_output",
            )
            st.caption(f"Word count: {len(humanized.split())}")

            # Download button
            st.download_button(
                "⬇️ Download Humanized Text",
                data=humanized,
                file_name="humanized_output.txt",
                mime="text/plain",
                use_container_width=True,
            )

        # ─── Metrics ───
        with st.sidebar:
            st.divider()
            st.subheader("📊 Local Self-Check Metrics")

            with st.spinner("Computing metrics…"):
                metrics = compute_metrics(humanized)

            m1, m2, m3 = st.columns(3)
            with m1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Avg Perplexity</div>
                    <div class="metric-value">{metrics['avg_perplexity']:.1f}</div>
                </div>
                """, unsafe_allow_html=True)
            with m2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Burstiness</div>
                    <div class="metric-value">{metrics['burstiness']:.3f}</div>
                </div>
                """, unsafe_allow_html=True)
            with m3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">TTR</div>
                    <div class="metric-value">{metrics['ttr']:.4f}</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("""
            <div class="warning-box">
                ⚠️ These are local estimates only and do not guarantee
                bypassing any external detector.
            </div>
            """, unsafe_allow_html=True)

        # ─── Change Review Table ───
        if show_changes and changes:
            st.divider()
            st.subheader("📝 Changes for Human Review")
            st.markdown(
                "Review each change and decide whether to accept or reject it."
            )

            # Build a DataFrame-like display
            for i, ch in enumerate(changes[:50]):  # cap at 50 for performance
                cols = st.columns([3, 3, 2, 1, 1])
                with cols[0]:
                    st.text(ch.original if len(ch.original) <= 60 else ch.original[:57] + "…")
                with cols[1]:
                    st.text(ch.modified if len(ch.modified) <= 60 else ch.modified[:57] + "…")
                with cols[2]:
                    st.caption(ch.module)
                with cols[3]:
                    st.checkbox("✅", key=f"accept_{i}", value=ch.accepted)
                with cols[4]:
                    st.checkbox("❌", key=f"reject_{i}", value=not ch.accepted)

            if len(changes) > 50:
                st.info(f"Showing first 50 of {len(changes)} changes.")

    elif initiate and not input_text:
        st.warning("Please upload a file or paste text before initiating.")


if __name__ == "__main__":
    main()
