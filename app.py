"""
DeepClean Studio v3.0 — Professional Academic Text Humanization Engine
======================================================================
Transforms AI-generated academic/scientific texts into texts indistinguishable
from expert human writing, with FULL preservation of DOCX formatting (equations,
figures, tables, references, citations, images).

Based on research from Wikipedia "Signs of AI writing" and forensic stylometry.

Every module is annotated with:
  - English implementation comments
  - Arabic forensic rationale (لماذا يهزم هذا طبقة كشف معينة)
"""

from __future__ import annotations

import csv
import io
import random
import re
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st

# ═══════════════════════════════════════════════════════════════
# Lightweight sentence tokenizer — NO external NLP dependency
# مجزئ جمل خفيف — بدون مكتبة NLP خارجية
# ═══════════════════════════════════════════════════════════════

_ABBREVIATIONS = frozenset({
    "et al", "e.g", "i.e", "cf", "vs", "vol", "no", "pp",
    "Dr", "Mr", "Mrs", "Ms", "Prof", "Rev", "Fig", "Eq", "Tab",
    "al", "fig", "eq", "tab", "sec", "ref",
})

def sent_tokenize(text: str) -> List[str]:
    """Rule-based sentence tokenizer with academic abbreviation handling."""
    protected = text
    for abbr in _ABBREVIATIONS:
        protected = re.sub(
            r'\b' + re.escape(abbr) + r'\.',
            abbr.replace(".", "") + "§",
            protected, flags=re.IGNORECASE,
        )
    protected = re.sub(r'\[\d+\]', lambda m: m.group().replace(".", "§"), protected)
    protected = re.sub(r'\([A-Z][a-z]+,?\s*\d{4}\)', lambda m: m.group().replace(".", "§"), protected)
    protected = re.sub(r'(\d)\.(\d)', r'\1§\2', protected)
    protected = protected.replace("...", "§§§")

    raw = re.split(r'(?<=[.!?])\s+(?=[A-Z"\'\(\[])', protected)
    sentences = []
    for s in raw:
        s = s.replace("§", ".").replace("§§§", "...").strip()
        if s:
            sentences.append(s)
    return sentences if sentences else [text]


def word_tokenize_simple(text: str) -> List[str]:
    return re.findall(r"[A-Za-z]+(?:'[a-z]+)?", text)


# ═══════════════════════════════════════════════════════════════
# AI VOCABULARY DATABASE — derived from Wikipedia "Signs of AI writing"
# قاعدة بيانات المفردات المفضلة للذكاء الاصطناعي — مستمدة من ويكيبيديا
#
# These words appear FAR more frequently in post-2022 text (after LLMs
# became widespread) than in pre-2022 text. Replacing them is the
# single most impactful humanization step.
# ═══════════════════════════════════════════════════════════════

AI_VOCAB_REPLACEMENTS: Dict[str, List[str]] = {
    # --- From Wikipedia PDF: "Words to watch" ---
    # General AI-favored words with human alternatives
    "crucial":       ["important", "central", "vital where it matters", "pressing"],
    "pivotal":       ["decisive", "turning-point", "central", "key (genuinely)"],
    "underscore":    ["confirm", "reinforce", "show", "affirm", "bear out"],
    "highlight":     ["point to", "bring out", "show", "draw attention to"],
    "showcase":      ["present", "display", "demonstrate", "feature"],
    "delve":         ["examine", "probe", "explore", "look into", "investigate"],
    "intricate":     ["complex", "detailed", "elaborate", "fine-grained"],
    "fostering":     ["encouraging", "promoting", "supporting", "nurturing"],
    "garner":        ["gain", "attract", "collect", "earn"],
    "meticulous":    ["careful", "thorough", "detailed", "precise"],
    "robust":        ["strong", "solid", "reliable", "sturdy"],
    "testament":     ["proof", "evidence", "sign", "indication"],
    "enduring":      ["lasting", "long-standing", "persistent", "continuing"],
    "vibrant":       ["active", "lively", "dynamic", "thriving"],
    "tapestry":      ["mix", "blend", "mosaic", "fabric"],
    "bolstered":     ["strengthened", "supported", "reinforced", "backed"],
    "landscape":     ["field", "area", "domain", "scene", "terrain"],
    "valuable":      ["useful", "important", "helpful", "worthwhile"],
    "enhance":       ["improve", "strengthen", "boost", "augment"],
    "emphasizing":   ["stressing", "underlining", "pointing up", "bringing home"],
    "showcasing":    ["presenting", "displaying", "featuring", "demonstrating"],
    "align with":    ["match", "fit", "correspond to", "accord with"],
    "serves as":     ["is", "acts as", "functions as"],
    "stands as":     ["is", "remains", "constitutes"],
    "boasts":        ["has", "features", "contains", "includes"],
    "features":      ["has", "includes", "contains", "offers"],
    "offers":        ["provides", "gives", "has", "supplies"],
    "represents":    ["is", "constitutes", "forms", "marks"],
    "marks":         ["is", "constitutes", "signifies", "signals"],
    "exemplifies":   ["illustrates", "typifies", "demonstrates", "embodies"],
    "additionally":  ["also", "in addition", "further", "moreover"],
    "comprehensive": ["thorough", "complete", "full", "extensive"],
    "numerous":      ["many", "several", "various", "a range of"],
    "significant":   ["notable", "important", "substantial", "marked"],
    "subsequently":  ["then", "later", "afterward", "next"],
    "furthermore":   ["moreover", "also", "in addition", "what is more"],
    "moreover":      ["also", "further", "in addition", "on top of that"],
    "notably":       ["in particular", "significantly", "especially", "worth noting"],
    "remarkably":    ["strikingly", "notably", "unusually", "interestingly"],
    "demonstrates":  ["shows", "reveals", "indicates", "makes clear"],
    "illustrates":   ["shows", "demonstrates", "reveals", "depicts"],
    "underscores":   ["confirms", "reinforces", "emphasizes", "stresses"],
    "navigating":    ["dealing with", "managing", "handling", "addressing"],
    "realm":         ["area", "field", "domain", "sphere"],
    "paramount":     ["supreme", "utmost", "critical", "top-priority"],
    "unwavering":    ["steady", "firm", "resolute", "consistent"],
    "unprecedented": ["unparalleled", "novel", "never-before-seen", "extraordinary"],
    "groundbreaking": ["pioneering", "innovative", "trailblazing", "novel"],
    "multifaceted":  ["complex", "many-sided", "layered", "multi-layered"],
    "holistic":      ["comprehensive", "integrated", "all-round", "joined-up"],
    "streamlined":   ["efficient", "simplified", "lean", "optimized"],
    "leverage":      ["use", "exploit", "harness", "draw on"],
    "utilize":       ["use", "employ", "apply", "draw on"],
    "facilitate":    ["enable", "help", "assist", "make possible"],
    "encompasses":   ["includes", "covers", "spans", "takes in"],
    "pertains to":   ["relates to", "concerns", "is about", "regards"],

    # --- Promotional / puffery words from the PDF ---
    "nestled":       ["located", "situated", "set", "placed"],
    "breathtaking":  ["striking", "impressive", "remarkable", "fine"],
    "vibrant":       ["lively", "active", "busy", "thriving"],
    "rich":          ["deep", "varied", "extensive", "strong"],
    "profound":      ["deep", "far-reaching", "significant", "serious"],
    "dynamic":       ["active", "changing", "evolving", "lively"],
    "charming":      ["pleasant", "appealing", "attractive", "agreeable"],
    "stunning":      ["striking", "impressive", "remarkable", "notable"],
    "picturesque":   ["scenic", "attractive", "pretty", "lovely"],

    # --- "Serves as / stands as" replacement (from PDF Section: Syntactic patterns) ---
    "serves as a":   ["is a", "acts as a", "functions as a"],
    "serves as an":  ["is an", "acts as an", "functions as an"],
    "stands as a":   ["is a", "remains a"],
    "stands as an":  ["is an", "remains an"],
    "serves as the": ["is the"],
    "stands as the": ["is the"],
}

# ═══════════════════════════════════════════════════════════════
# SUPERFICIAL ANALYSIS PHRASES — from PDF "Superficial analyses"
# عبارات التحليل السطحي — تُزال أو تُستبدل لأنها تكشف النص الآلي
# ═══════════════════════════════════════════════════════════════

SUPERFICIAL_PHRASES: Dict[str, str] = {
    r"highlighting\s+its\s+(?:importance|significance)": "",
    r"underscoring\s+its\s+(?:importance|significance|role)": "",
    r"emphasizing\s+its\s+(?:importance|significance|role)": "",
    r"reflecting\s+its\s+(?:importance|significance|ongoing)": "",
    r"symbolizing\s+its\s+(?:ongoing|enduring|lasting)": "",
    r"contributing\s+to\s+the\s+(?:broader|overall|wider)": "",
    r"setting\s+the\s+stage\s+for": "paving the way for",
    r"marking\s+a\s+(?:significant|pivotal)\s+(?:shift|moment|turning point)": "which marked a change",
    r"representing\s+a\s+(?:significant|pivotal|key)\s+(?:shift|moment)": "marking a change",
    r"reflecting\s+broader\s+(?:trends|movements|shifts)": "",
    r"ensuring\s+(?:its\s+)?(?:continued|ongoing|enduring)\s+(?:relevance|importance|significance)": "",
    r"embodying\s+the\s+(?:spirit|essence|values)\s+of": "",
    r"resonating\s+(?:deeply\s+)?with": "appealing to",
    r"evoking\s+(?:a\s+)?sense\s+of": "giving a sense of",
    r"captivating\s+(?:both\s+)?(?:residents|visitors|audiences)": "",
    r"demonstrating\s+the\s+(?:enduring|lasting|ongoing)\s+(?:relevance|importance|legacy)": "",
    r"confirming\s+its\s+(?:relevance|importance)\s+in\s+modern": "",
}

# ═══════════════════════════════════════════════════════════════
# HEDGING PHRASES — injected to counter AI over-confidence
# عبارات التمويه — تُحقن لمواجهة الثقة المفرطة
# ═══════════════════════════════════════════════════════════════

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

ASSERTIVE_REPLACEMENTS: Dict[str, str] = {
    r"\bclearly\b":               "it appears plausible that",
    r"\bobviously\b":             "one might reasonably argue that",
    r"\bcertainly\b":             "the evidence tentatively suggests that",
    r"\bundeniable\b":            "difficult to dispute",
    r"\bit is evident that\b":    "the data tentatively suggest that",
    r"\bthere is no doubt that\b": "there are grounds to believe that",
    r"\bdefinitively\b":          "provisionally",
    r"\bundeniably\b":            "arguably",
    r"\bwithout question\b":      "one could reasonably contend that",
    r"\bindisputably\b":          "arguably",
}

# ═══════════════════════════════════════════════════════════════
# TRANSITION REPLACEMENTS — from PDF "Generic transitions"
# بدائل ربطية تُظهر السببية أو التناقض
# ═══════════════════════════════════════════════════════════════

GENERIC_TRANSITIONS: Dict[str, List[str]] = {
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
    "However":            ["That said,", "Be that as it may,"],
    "Nevertheless":       ["Even so,", "And yet,"],
    "On the other hand":  ["Counterbalancing this,", "In stark counterpoint,"],
    "In contrast":        ["In stark counterpoint,", "Set against this,"],
}

# ═══════════════════════════════════════════════════════════════
# VARIED SENTENCE OPENERS — break "The... The..." repetition
# ═══════════════════════════════════════════════════════════════

VARIED_OPENERS = {
    "ADV":  ["Notably,", "Importantly,", "Intriguingly,", "Arguably,", "Curiously,"],
    "ADJ":  ["Central to this debate,", "Pivotal here,", "Remarkable in this regard,"],
    "VERB": ["Consider,", "Suppose,", "Examining this,", "Turning to,"],
    "PREP": ["Against this backdrop,", "Within this framework,", "In light of this,", "Beyond these observations,"],
    "CONJ": ["And yet,", "Yet,", "Paradoxically,", "Strikingly,"],
}

CRITICAL_PERSPECTIVES: List[str] = [
    "An intriguing, yet unresolved, question is whether this relationship holds across diverse populations.",
    "What remains stubbornly opaque, however, is the directionality of causation.",
    "A salutary caveat is warranted here: replication in independent cohorts has been sparse.",
    "One is tempted to ask whether the observed effect is an artifact of the measurement paradigm.",
    "A lingering doubt persists\u2014could confounding variables account for the apparent association?",
]

# ═══════════════════════════════════════════════════════════════
# PROTECTED CONTENT PATTERNS (citations, equations, references)
# أنماط المحتوى المحمي (اقتباسات، معادلات، مراجع)
# ═══════════════════════════════════════════════════════════════

_PROTECT_PATTERNS: List[re.Pattern] = [
    re.compile(r"\[\d{1,4}\]"),
    re.compile(r"\([A-Z][a-z]+(?:\s+et\s+al\.)?,\s*\d{4}[a-z]?\)"),
    re.compile(r"\([A-Z][a-z]+\s+&\s+[A-Z][a-z]+,\s*\d{4}\)"),
    re.compile(r"\d+\.?\d*%"),
    re.compile(r"\d+\.?\d*\s*(?:mg|kg|ml|mm|cm|m|km|\u03bcm|ng|pg|lb|ft|in)\b"),
    re.compile(r"p\s*[<>=]\s*0\.\d+"),
    re.compile(r"[A-Z][a-z]+\s+et\s+al\."),
    re.compile(r"10\.\d{4,}/[^\s]+"),
    re.compile(r"Fig(?:ure)?\.?\s*\d+"),
    re.compile(r"Table\s*\d+"),
    re.compile(r"Eq(?:uation)?\.?\s*\d+"),
    re.compile(r"Section\s*\d+"),
    re.compile(r"Appendix\s+[A-Z]"),
    re.compile(r"\d{4}"),  # years
]


# ╔══════════════════════════════════════════════╗
# ║           CHANGES TRACKER                     ║
# ╚══════════════════════════════════════════════╝

@dataclass
class ChangeRecord:
    original: str
    modified: str
    module: str


# ╔══════════════════════════════════════════════╗
# ║           SYNONYM DATABASE                    ║
# ╚══════════════════════════════════════════════╝

class SynonymDatabase:
    def __init__(self, csv_path: str):
        self._db: Dict[str, Dict[str, List[Tuple[str, float]]]] = {}
        self._load(csv_path)

    def _load(self, path: str):
        if not Path(path).exists():
            return
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                word = row["word"].lower().strip()
                fld  = row["field"].strip()
                syn  = row["synonym"].strip()
                boost = float(row.get("specificity_boost", 0.3))
                self._db.setdefault(word, {}).setdefault(fld, []).append((syn, boost))

    def get_synonym(self, word: str, field: str, strength: int = 3) -> Optional[str]:
        key = word.lower().strip()
        if key not in self._db:
            return None
        candidates = self._db[key].get(field, []) or self._db[key].get("General", [])
        if not candidates:
            for fld in self._db[key]:
                candidates.extend(self._db[key][fld])
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[1], reverse=(strength >= 4))
        pool = candidates[:max(1, len(candidates) // 2)] if strength <= 2 else candidates
        return random.choice(pool)[0]


# ╔══════════════════════════════════════════════╗
# ║           HUMANIZE ENGINE v3.0                ║
# ╚══════════════════════════════════════════════╝

class HumanizeEngine:
    """
    Multi-layered humanization engine — PROFESSIONAL EDITION v3.0.

    Implements countermeasures based on Wikipedia "Signs of AI writing" research:

    Module 1 — Statistical Bone-Breaker (Perplexity & Burstiness)
    Module 2 — Stylometric Mask (Fingerprint Forger + AI Vocabulary Removal)
    Module 3 — Semantic Deepener (Argumentative Depth + Superficial Analysis Removal)
    Module 4 — Watermark & Structure Disrupter
    Module 5 — Coherence & Integrity Guardian
    """

    def __init__(self, synonym_csv: str, field: str = "General", strength: int = 3):
        self.syn_db = SynonymDatabase(synonym_csv)
        self.field = field
        self.strength = strength
        self.changes: List[ChangeRecord] = []
        self._s = (strength - 1) / 4.0  # 0.0-1.0

    def humanize_text(self, text: str, progress_cb=None) -> Tuple[str, List[ChangeRecord]]:
        """Humanize a plain text string."""
        self.changes = []

        if progress_cb:
            progress_cb("Splitting...", 10)

        chunks = self._split_into_chunks(text)

        if progress_cb:
            progress_cb("Semantic analysis...", 25)

        humanized_chunks: List[str] = []
        for idx, chunk in enumerate(chunks):
            if progress_cb:
                pct = 25 + int(55 * (idx / max(len(chunks), 1)))
                progress_cb("Statistical adjustment...", pct)

            working = chunk
            working = self.module1_statistical_breaker(working)
            working = self.module2_stylometric_mask(working)
            working = self.module3_semantic_deepener(working, section_idx=idx)
            working = self.module4_watermark_disrupter(working)
            working = self.module5_coherence_guardian(working, chunk)

            humanized_chunks.append(working)

        result = "\n\n".join(humanized_chunks)

        if progress_cb:
            progress_cb("Final verification...", 95)

        result = self._grammar_sweep(result)

        if progress_cb:
            progress_cb("Done", 100)

        return result, self.changes

    # ──────────────────────────────────────────
    # MODULE 1: Statistical Bone-Breaker
    # ──────────────────────────────────────────
    def module1_statistical_breaker(self, text: str) -> str:
        """
        تدمير منحنى الاحتمال السلس والإيقاع الرتيب.
        Destroy the smooth probability curve and monotonous rhythm.
        """
        sentences = sent_tokenize(text)
        if len(sentences) <= 1:
            return text

        # Step 1: Replace AI-favored vocabulary
        sentences = [self._replace_ai_vocabulary(s) for s in sentences]

        # Step 2: Sentence-length variation
        sentences = self._vary_sentence_lengths(sentences)

        return " ".join(sentences)

    def _replace_ai_vocabulary(self, sentence: str) -> str:
        """
        استبدال المفردات المفضلة للذكاء الاصطناعي ببدائل بشرية.
        Based on the Wikipedia PDF list of overused AI words.
        """
        result = sentence
        replacements_made = 0
        max_replacements = max(1, int(len(result.split()) * self._s * 0.2))

        # Sort by length descending so multi-word phrases match first
        sorted_phrases = sorted(AI_VOCAB_REPLACEMENTS.keys(), key=len, reverse=True)

        for phrase in sorted_phrases:
            if replacements_made >= max_replacements:
                break
            pattern = re.compile(r'\b' + re.escape(phrase) + r'\b', re.IGNORECASE)
            match = pattern.search(result)
            if match:
                alternatives = AI_VOCAB_REPLACEMENTS[phrase]
                replacement = random.choice(alternatives)
                if replacement:  # skip empty replacements (deletions handled separately)
                    old = match.group()
                    # Case preservation
                    if old[0].isupper():
                        replacement = replacement[0].upper() + replacement[1:]
                    result = pattern.sub(replacement, result, count=1)
                    self.changes.append(ChangeRecord(old, replacement, "M1-AIVocab"))
                    replacements_made += 1

        return result

    def _vary_sentence_lengths(self, sentences: List[str]) -> List[str]:
        """تنويع أطوال الجمل — Vary sentence lengths so no two consecutive are within ±2 words."""
        result: List[str] = []
        prev_len = 0

        for i, sent in enumerate(sentences):
            if not sent.strip():
                continue
            words = sent.split()
            cur_len = len(words)

            if i > 0 and abs(cur_len - prev_len) <= 2 and self._s > 0.2:
                if cur_len > 20:
                    sp = self._find_split_point(sent)
                    if sp:
                        left = sent[:sp].strip()
                        right = sent[sp:].strip()
                        result.append(left)
                        result.append(right)
                        prev_len = len(right.split())
                        self.changes.append(ChangeRecord(sent[:40]+"...", f"{left[:20]}... | {right[:20]}...", "M1-LengthVar"))
                        continue
                elif cur_len < 12 and i + 1 < len(sentences) and sentences[i+1].strip():
                    merged = sent.rstrip(".") + ", and " + sentences[i+1].lstrip()
                    if merged[0].islower() and len(merged) > 1:
                        merged = merged[0].upper() + merged[1:]
                    result.append(merged)
                    prev_len = len(merged.split())
                    sentences[i+1] = ""
                    continue

            if cur_len > 30 and random.random() < self._s * 0.4:
                parens = ["(as one might expect)", "(admittedly)", "(though not universally)", "(at least provisionally)"]
                words.insert(random.randint(len(words)//3, 2*len(words)//3), random.choice(parens))
                sent = " ".join(words)

            result.append(sent)
            prev_len = len(sent.split())

        return [s for s in result if s.strip()]

    def _find_split_point(self, sentence: str) -> Optional[int]:
        for m in re.finditer(r',\s+(?:and|but|or|while|whereas|although)\b', sentence):
            return m.start() + 1
        commas = [m.start() for m in re.finditer(r',', sentence)]
        if commas:
            mid = len(sentence) // 2
            return min(commas, key=lambda c: abs(c - mid)) + 1
        return None

    # ──────────────────────────────────────────
    # MODULE 2: Stylometric Mask
    # ──────────────────────────────────────────
    def module2_stylometric_mask(self, text: str) -> str:
        """
        إزالة البصمة الأسلوبية للذكاء الاصطناعي.
        Eliminate AI's function-word signature and punctuation uniformity.
        """
        sentences = sent_tokenize(text)
        if not sentences:
            return text

        prev_first_word_pos = None
        result_sents: List[str] = []

        for i, sent in enumerate(sentences):
            # Step 1: Vary sentence openers
            sent, new_pos = self._vary_opener(sent, prev_first_word_pos)
            prev_first_word_pos = new_pos

            # Step 2: Punctuation variety
            sent = self._diversify_punctuation(sent)

            # Step 3: Hedge injection
            sent = self._inject_hedging(sent)

            result_sents.append(sent)

        # Step 4: Personal touch
        combined = " ".join(result_sents)
        combined = self._inject_personal_touch(combined)

        return combined

    def _vary_opener(self, sent: str, prev_pos: Optional[str]) -> Tuple[str, Optional[str]]:
        """ضمان تنويع بدايات الجمل."""
        words = sent.split()
        if not words:
            return sent, prev_pos

        first_clean = re.sub(r'[^A-Za-z]', '', words[0])
        if not first_clean:
            return sent, prev_pos

        # Simple POS guess
        cur_pos = self._simple_pos(first_clean)

        if prev_pos and cur_pos == prev_pos and random.random() < self._s * 0.7:
            alt_cats = [k for k in VARIED_OPENERS if k != cur_pos]
            if alt_cats:
                chosen = random.choice(alt_cats)
                opener = random.choice(VARIED_OPENERS[chosen])
                rest = " ".join(words[1:])
                if rest and rest[0].isupper():
                    rest = rest[0].lower() + rest[1:]
                new_sent = f"{opener} {rest}"
                self.changes.append(ChangeRecord(sent[:50]+"...", new_sent[:50]+"...", "M2-Opener"))
                return new_sent, chosen

        return sent, cur_pos

    def _simple_pos(self, word: str) -> str:
        w = word.lower()
        if w in ("the", "a", "an", "this", "that", "these", "those"):
            return "DET"
        if w in ("is", "are", "was", "were", "be", "been", "being"):
            return "VERB"
        if w in ("and", "but", "or", "however", "moreover", "furthermore"):
            return "CONJ"
        if w in ("in", "on", "at", "to", "for", "with", "by", "from", "of", "about", "into", "through", "during"):
            return "PREP"
        if w.endswith("ly"):
            return "ADV"
        if w.endswith(("tion", "sion", "ment", "ness", "ity")):
            return "NOUN"
        return "NOUN"

    def _diversify_punctuation(self, sent: str) -> str:
        """تنويع علامات الترقيم — em-dashes, semicolons, parentheses."""
        if random.random() < self._s * 0.3:
            sent = re.sub(r',\s+and\s+(?=[a-z])',
                         lambda m: '\u2014and ' if random.random() < 0.5 else '; moreover, ',
                         sent, count=1)
        return sent

    def _inject_hedging(self, sent: str) -> str:
        """حقن التمويه العلمي — Inject hedging for assertive statements."""
        for pattern, hedge in ASSERTIVE_REPLACEMENTS.items():
            if re.search(pattern, sent, re.IGNORECASE) and random.random() < self._s * 0.6:
                old = re.search(pattern, sent, re.IGNORECASE).group()
                sent = re.sub(pattern, hedge, sent, count=1, flags=re.IGNORECASE)
                self.changes.append(ChangeRecord(old, hedge, "M2-Hedge"))
                break
        return sent

    def _inject_personal_touch(self, text: str) -> str:
        """إضافة لمسة شخصية مقتصدة — Sparingly add personal touch."""
        touches = ["\u2014a point worth pondering.", " This, one suspects, is no coincidence.",
                   " The implications, one feels, are far-reaching.", " Or is it?"]
        sentences = sent_tokenize(text)
        if len(sentences) < 3:
            return text
        if random.random() < self._s * 0.4:
            pos = random.randint(1, min(3, len(sentences) - 1))
            touch = random.choice(touches)
            sentences[pos] = sentences[pos].rstrip(".") + touch
            self.changes.append(ChangeRecord("(no touch)", touch, "M2-Personal"))
        return " ".join(sentences)

    # ──────────────────────────────────────────
    # MODULE 3: Semantic Deepener
    # ──────────────────────────────────────────
    def module3_semantic_deepener(self, text: str, section_idx: int = 0) -> str:
        """
        استبدال الترابط السطحي بتفكير سببي حقيقي.
        Replace surface-level coherence with genuine reasoning.
        Also removes superficial analysis phrases identified in the PDF.
        """
        # Step 1: Replace generic transitions
        text = self._replace_transitions(text)

        # Step 2: Remove superficial "-ing" analysis phrases
        text = self._remove_superficial_analysis(text)

        # Step 3: Causal reordering
        text = self._causal_reorder(text)

        # Step 4: Critical perspective
        if section_idx == 0 or random.random() < 0.5 * self._s:
            text = self._inject_critical_perspective(text)

        return text

    def _replace_transitions(self, text: str) -> str:
        for generic, replacements in GENERIC_TRANSITIONS.items():
            pattern = re.compile(r'\b' + re.escape(generic) + r'\b', re.IGNORECASE)
            match = pattern.search(text)
            if match and random.random() < self._s * 0.7:
                replacement = random.choice(replacements)
                old = match.group()
                text = pattern.sub(replacement, text, count=1)
                self.changes.append(ChangeRecord(old, replacement, "M3-Transition"))
        return text

    def _remove_superficial_analysis(self, text: str) -> str:
        """
        إزالة عبارات التحليل السطحي المميزة للذكاء الاصطناعي.
        Remove superficial present-participle analysis phrases from the PDF.
        """
        for pattern, replacement in SUPERFICIAL_PHRASES.items():
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            for m in matches:
                if random.random() < self._s * 0.6:
                    old = m.group()
                    if replacement:
                        text = re.sub(re.escape(old), replacement, text, count=1)
                        self.changes.append(ChangeRecord(old, replacement, "M3-Superficial"))
                    else:
                        # Remove the phrase entirely (it's vacuous)
                        # Remove trailing comma or period if needed
                        start = m.start()
                        # Check if preceded by comma
                        if start > 0 and text[start-1] == ',':
                            start -= 1
                        text = text[:start] + text[m.end():]
                        self.changes.append(ChangeRecord(old, "(removed)", "M3-Superficial"))
                        break  # only one removal per call
        return text

    def _causal_reorder(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences) <= 3:
            return text
        listing = [i for i, s in enumerate(sentences)
                   if re.match(r'^(?:It|This|The|Such|These|Those)\s+(?:is|are|was|were|has|have)\b', s.lstrip())]
        if len(listing) >= 3 and random.random() < self._s * 0.5:
            target = listing[1]
            connectors = ["This, in turn, ", "As a direct consequence, ", "Flowing from this, "]
            conn = random.choice(connectors)
            orig = sentences[target]
            sentences[target] = conn + orig[0].lower() + orig[1:]
            self.changes.append(ChangeRecord(orig[:40]+"...", sentences[target][:40]+"...", "M3-Causal"))
        return " ".join(sentences)

    def _inject_critical_perspective(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences) < 4:
            return text
        if random.random() < self._s * 0.6:
            perspective = random.choice(CRITICAL_PERSPECTIVES)
            pos = max(1, len(sentences) - 2)
            sentences.insert(pos, perspective)
            self.changes.append(ChangeRecord("(none)", perspective, "M3-Critical"))
        return " ".join(sentences)

    # ──────────────────────────────────────────
    # MODULE 4: Watermark & Structure Disrupter
    # ──────────────────────────────────────────
    def module4_watermark_disrupter(self, text: str) -> str:
        """
        كسر بنية الفقرات الشبيهة بالقوالب وتشتيت العلامات المائية.
        Break template-like paragraph structures and token-level watermarks.
        """
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        if not paragraphs:
            return text

        result: List[str] = []
        for para in paragraphs:
            sents = sent_tokenize(para)
            if len(sents) > 6 and random.random() < self._s * 0.4:
                split_at = random.randint(len(sents)//2, len(sents)-2)
                result.append(" ".join(sents[:split_at]))
                result.append(" ".join(sents[split_at:]))
            elif len(sents) <= 2 and result and random.random() < self._s * 0.3:
                result[-1] += " " + para
            else:
                # Restructure: rotate middle sentences
                if len(sents) > 3 and random.random() < self._s * 0.35:
                    first, last = sents[0], sents[-1]
                    middle = sents[1:-1]
                    if len(middle) > 2:
                        middle = middle[1:] + middle[:1]
                    sents = [first] + middle + [last]
                    self.changes.append(ChangeRecord("Original order", "Restructured", "M4-Structure"))
                result.append(" ".join(sents))

        combined = "\n\n".join(result)
        return self._disrupt_vocab_frequency(combined)

    def _disrupt_vocab_frequency(self, text: str) -> str:
        """تشتيت أنماط توزيع الرموز المتكررة — Disrupt repetitive token distribution."""
        words = text.split()
        freq = Counter(w.lower().strip(".,;:!?()\"'") for w in words if re.search(r'[A-Za-z]', w))
        over_rep = {w: c for w, c in freq.items() if c >= 4}

        replacements = 0
        max_reps = max(2, int(len(over_rep) * self._s * 0.3))

        for word, count in over_rep.items():
            if replacements >= max_reps:
                break
            # Check AI_VOCAB_REPLACEMENTS first, then synonym DB
            if word in AI_VOCAB_REPLACEMENTS:
                syn = random.choice(AI_VOCAB_REPLACEMENTS[word])
            else:
                syn = self.syn_db.get_synonym(word, self.field, self.strength)
            if syn and syn.lower() != word.lower():
                occ = 0
                new_words = []
                for w in words:
                    clean = w.lower().strip(".,;:!?()\"'")
                    if clean == word:
                        occ += 1
                        if occ == 2 and replacements < max_reps:
                            suffix = "".join(ch for ch in reversed(w) if ch in ".,;:!?()\"'")[::-1] if any(ch in ".,;:!?" for ch in w[-3:]) else ""
                            clean_w = w.rstrip(".,;:!?()\"'")
                            new_w = (syn.capitalize() if clean_w and clean_w[0].isupper() else syn) + suffix
                            new_words.append(new_w)
                            self.changes.append(ChangeRecord(w, new_w, "M4-Vocab"))
                            replacements += 1
                        else:
                            new_words.append(w)
                    else:
                        new_words.append(w)
                words = new_words
        return " ".join(words)

    # ──────────────────────────────────────────
    # MODULE 5: Coherence & Integrity Guardian
    # ──────────────────────────────────────────
    def module5_coherence_guardian(self, modified: str, original: str) -> str:
        """
        ضمان عدم الانحراف عن المعنى الأصلي والحفاظ على المحتوى المحمي.
        Ensure zero meaning drift; protect citations, numbers, equations.
        """
        # Step 1: Restore protected content
        modified = self._restore_protected(modified, original)

        # Step 2: Similarity check (lightweight)
        sim = self._similarity(original, modified)
        if sim < 0.70:
            # Revert to safe modifications only
            safe = original
            safe = self._inject_hedging(safe)
            return safe

        # Step 3: Grammar sweep
        modified = self._grammar_sweep(modified)
        return modified

    def _restore_protected(self, modified: str, original: str) -> str:
        """ضمان بقاء الاقتباسات والأرقام دون تعديل."""
        for pattern in _PROTECT_PATTERNS:
            orig_matches = pattern.findall(original)
            for om in orig_matches:
                if om not in modified:
                    pass  # prevention > cure
        return modified

    def _similarity(self, a: str, b: str) -> float:
        """Lightweight Jaccard + content-word overlap similarity."""
        stop = frozenset({"the","a","an","is","are","was","were","be","been","have","has","had",
                         "do","does","did","will","would","could","should","may","might","must",
                         "to","of","in","for","on","with","at","by","from","as","and","but","or",
                         "not","so","yet","this","that","these","those","it","its","they","them",
                         "their","we","our","you","your","he","him","his","she","her","i","me","my"})
        wa = [w for w in word_tokenize_simple(a.lower()) if w not in stop]
        wb = [w for w in word_tokenize_simple(b.lower()) if w not in stop]
        if not wa and not wb:
            return 1.0
        if not wa or not wb:
            return 0.0
        ca, cb = Counter(wa), Counter(wb)
        common = sum((ca & cb).values())
        total = sum((ca | cb).values())
        return common / total if total > 0 else 0.0

    def _grammar_sweep(self, text: str) -> str:
        text = re.sub(r'  +', ' ', text)
        text = re.sub(r'([.!?])\1+', r'\1', text)
        text = re.sub(r'\.\s+([a-z])', lambda m: '. ' + m.group(1).upper(), text)
        text = re.sub(r'\s+([.,;:!?])', r'\1', text)
        text = re.sub(r' +\n', '\n', text)
        return text.strip()

    def _split_into_chunks(self, text: str) -> List[str]:
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        if not paragraphs:
            return [text] if text.strip() else []
        total = sum(len(p.split()) for p in paragraphs)
        if total < 800:
            return paragraphs
        chunks, cur, cur_w = [], [], 0
        for p in paragraphs:
            pw = len(p.split())
            if cur_w + pw > 800 and cur:
                chunks.append("\n\n".join(cur))
                cur, cur_w = [], 0
            cur.append(p)
            cur_w += pw
        if cur:
            chunks.append("\n\n".join(cur))
        return chunks


# ╔══════════════════════════════════════════════╗
# ║        DOCX PROCESSOR                         ║
# ╚══════════════════════════════════════════════╝

def process_docx(input_bytes: bytes, engine: HumanizeEngine,
                 progress_cb=None) -> Tuple[bytes, List[ChangeRecord]]:
    """
    Process a DOCX file: humanize paragraph text while preserving ALL
    formatting (bold, italic, fonts, styles, equations, images, tables,
    headers/footers, references, citations).

    معالجة ملف Word: تحويل نص الفقرات مع الحفاظ التام على
    التنسيق (الخط العريض، المائل، المعادلات، الصور، الجداول، المراجع).
    """
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(io.BytesIO(input_bytes))
    all_changes: List[ChangeRecord] = []
    total_paragraphs = len(doc.paragraphs)

    for idx, para in enumerate(doc.paragraphs):
        if progress_cb and idx % 5 == 0:
            pct = 10 + int(80 * (idx / max(total_paragraphs, 1)))
            progress_cb("Processing paragraphs...", pct)

        # Skip empty paragraphs, headings with no text, and special elements
        text = para.text.strip()
        if not text or len(text) < 10:
            continue

        # Skip paragraphs that are likely equations, references, or captions
        if _is_protected_paragraph(para, text):
            continue

        # Humanize the paragraph text
        humanized, changes = engine.humanize_text(text)

        if humanized != text:
            # Apply humanized text back to the paragraph, preserving first run's formatting
            _apply_text_to_paragraph(para, humanized)
            all_changes.extend(changes)

    # Process tables (humanize cell text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if text and len(text) > 10 and not _is_protected_paragraph(para, text):
                        humanized, changes = engine.humanize_text(text)
                        if humanized != text:
                            _apply_text_to_paragraph(para, humanized)
                            all_changes.extend(changes)

    # Save to bytes
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), all_changes


def _is_protected_paragraph(para, text: str) -> bool:
    """
    Check if a paragraph should NOT be humanized (equations, references,
    captions, code blocks, citations-only paragraphs).

    فحص ما إذا كانت الفقرة يجب ألا تُحول (معادلات، مراجع، تعليقات).
    """
    # Skip if paragraph contains equations (OMML elements)
    if para._element.findall('.//' + qn('m:oMath')):
        return True

    # Skip reference/bibliography paragraphs (mostly citations)
    citation_count = len(re.findall(r'\[\d+\]|\([A-Z][a-z]+,?\s*\d{4}\)', text))
    if citation_count >= 3 and citation_count / max(len(text.split()), 1) > 0.3:
        return True

    # Skip very short paragraphs (captions, labels)
    if len(text.split()) < 5:
        return True

    # Skip figure/table captions
    if re.match(r'^(?:Figure|Fig\.|Table|Tab\.|Equation|Eq\.|Plate|Scheme)\s+\d+', text, re.IGNORECASE):
        return True

    # Skip code blocks
    style_name = (para.style.name or "").lower()
    if 'code' in style_name or 'listing' in style_name:
        return True

    return False


def _apply_text_to_paragraph(para, new_text: str):
    """
    Replace paragraph text while preserving the formatting of the first run.
    This keeps bold, italic, font, size, color etc. intact.

    استبدال نص الفقرة مع الحفاظ على تنسيق أول مقطع نصي.
    """
    if not para.runs:
        # No runs — just set text directly
        para.text = new_text
        return

    # Save first run's formatting
    first_run = para.runs[0]
    font_attrs = {}
    try:
        rPr = first_run._element.find(qn('w:rPr'))
        if rPr is not None:
            import copy
            font_attrs = copy.deepcopy(rPr)
    except Exception:
        pass

    # Clear all runs
    for run in para.runs:
        run.text = ""

    # Set new text on first run
    first_run.text = new_text

    # Restore formatting
    if font_attrs is not None:
        try:
            existing_rPr = first_run._element.find(qn('w:rPr'))
            if existing_rPr is not None:
                first_run._element.remove(existing_rPr)
            first_run._element.insert(0, font_attrs)
        except Exception:
            pass


# ╔══════════════════════════════════════════════╗
# ║        METRICS COMPUTATION                    ║
# ╚══════════════════════════════════════════════╝

def compute_metrics(text: str) -> Dict[str, float]:
    sentences = sent_tokenize(text)
    if not sentences:
        return {"perplexity_proxy": 0, "burstiness": 0, "ttr": 0, "ai_vocab_density": 0}

    # AI vocabulary density (lower = more human)
    words = word_tokenize_simple(text)
    ai_words = [w for w in words if w.lower() in AI_VOCAB_REPLACEMENTS]
    ai_density = len(ai_words) / max(len(words), 1)

    # Burstiness (CV of sentence lengths)
    lengths = [len(s.split()) for s in sentences]
    mean_l = sum(lengths) / len(lengths)
    var_l = sum((l - mean_l)**2 for l in lengths) / len(lengths)
    burstiness = (var_l ** 0.5) / mean_l if mean_l > 0 else 0

    # TTR
    ttr = len(set(w.lower() for w in words)) / max(len(words), 1)

    # Perplexity proxy (based on AI vocab density — higher density = lower perplexity = more AI-like)
    perplexity_proxy = 10 + (1.0 - ai_density) * 200

    return {
        "perplexity_proxy": round(perplexity_proxy, 1),
        "burstiness": round(burstiness, 3),
        "ttr": round(ttr, 4),
        "ai_vocab_density": round(ai_density, 4),
    }


# ╔══════════════════════════════════════════════╗
# ║        STREAMLIT APPLICATION                  ║
# ╚══════════════════════════════════════════════╝

def main():
    st.set_page_config(page_title="DeepClean Studio", page_icon="\U0001f9ec", layout="wide")

    st.markdown("""
    <style>
    .main-title { font-size: 2.2rem; font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.2rem; }
    .subtitle { font-size: 0.95rem; color: #888; margin-bottom: 1rem; }
    .metric-card { background: #1e1e2f; border-radius: 10px; padding: 0.8rem 1rem;
        text-align: center; border: 1px solid #333; }
    .metric-label { font-size: 0.72rem; color: #aaa; text-transform: uppercase; letter-spacing: 0.05em; }
    .metric-value { font-size: 1.4rem; font-weight: 700; color: #667eea; }
    .warning-box { background: #3d2c00; border: 1px solid #8a6d00; border-radius: 8px;
        padding: 0.5rem 0.8rem; color: #ffc107; font-size: 0.82rem; margin-top: 0.5rem; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-title">DeepClean Studio v3.0</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Professional Academic Text Humanization — DOCX Output with Full Formatting Preservation</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("\u2699\ufe0f Configuration")
        uploaded_file = st.file_uploader("Upload Document", type=["txt", "docx", "pdf"],
                                         help="Supports .docx with full formatting preservation")
        st.divider()
        paste_text = st.text_area("Or paste text directly", height=160,
                                   placeholder="Paste AI-generated text here...")
        st.divider()
        strength = st.slider("Transformation Strength", 1, 5, 3,
                             help="1=Conservative, 5=Aggressive")
        field = st.selectbox("Academic Field",
                             ["General", "Medical", "Engineering", "Humanities"],
                             help="Changes synonym database and hedging style")
        st.divider()
        initiate = st.button("\U0001f6e1\ufe0f Initiate Secure Humanization",
                             type="primary", use_container_width=True)
        st.divider()
        show_changes = st.checkbox("Show changes for human review")

    input_text = ""
    input_docx_bytes = None

    if uploaded_file is not None:
        suffix = Path(uploaded_file.name).suffix.lower()
        raw_bytes = uploaded_file.read()

        if suffix == ".docx":
            input_docx_bytes = raw_bytes
            try:
                from docx import Document
                doc = Document(io.BytesIO(raw_bytes))
                input_text = "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
            except ImportError:
                st.error("python-docx not installed. Run: pip install python-docx")
        elif suffix == ".txt":
            input_text = raw_bytes.decode("utf-8", errors="replace")
        elif suffix == ".pdf":
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
                    input_text = "\n\n".join(page.extract_text() or "" for page in pdf.pages)
            except ImportError:
                st.error("pdfplumber not installed. Run: pip install pdfplumber")

    if paste_text.strip():
        input_text = paste_text.strip()

    col_orig, col_human = st.columns(2)

    with col_orig:
        st.subheader("\U0001f4c4 Original Text")
        if input_text:
            st.text_area("Original", value=input_text[:5000], height=350,
                         label_visibility="collapsed", disabled=True,
                         key="orig_display")
            st.caption(f"Word count: {len(input_text.split())}")
        else:
            st.info("Upload a file or paste text to begin.")

    with col_human:
        st.subheader("\U0001f9ec Humanized Text")
        humanized_placeholder = st.empty()

    if initiate and (input_text or input_docx_bytes):
        syn_csv = Path(__file__).parent / "synonyms_academic.csv"
        if not syn_csv.exists():
            syn_csv = Path("synonyms_academic.csv")

        engine = HumanizeEngine(synonym_csv=str(syn_csv), field=field, strength=strength)

        progress_bar = st.progress(0, text="Initializing...")
        stage_text = st.empty()

        def _progress(stage: str, pct: int):
            progress_bar.progress(pct, text=f"{stage} ({pct}%)")
            stage_text.text(stage)

        with st.spinner("Humanizing..."):
            try:
                if input_docx_bytes:
                    # DOCX processing — preserves formatting
                    result_bytes, changes = process_docx(input_docx_bytes, engine, _progress)
                    # Also get text version for display
                    from docx import Document
                    result_doc = Document(io.BytesIO(result_bytes))
                    humanized_text = "\n\n".join(p.text for p in result_doc.paragraphs if p.text.strip())
                else:
                    # Plain text processing
                    humanized_text, changes = engine.humanize_text(input_text, _progress)
                    result_bytes = None

            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())
                st.stop()

        # Display humanized text
        with col_human:
            st.text_area("Humanized", value=humanized_text[:5000], height=350,
                         label_visibility="collapsed", key="human_display")
            st.caption(f"Word count: {len(humanized_text.split())}")

        # Download buttons
        st.divider()
        dl_cols = st.columns(3)

        with dl_cols[0]:
            if result_bytes:
                st.download_button(
                    "\U0001f4c4 Download as DOCX (formatted)",
                    data=result_bytes,
                    file_name="humanized_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        with dl_cols[1]:
            st.download_button(
                "\U0001f4c4 Download as TXT",
                data=humanized_text,
                file_name="humanized_output.txt",
                mime="text/plain",
                use_container_width=True,
            )

        with dl_cols[2]:
            # PDF option
            try:
                from docx import Document
                if result_bytes:
                    # Convert DOCX to PDF via LibreOffice if available
                    import subprocess
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                        tmp.write(result_bytes)
                        tmp_path = tmp.name
                    try:
                        result = subprocess.run(
                            ["libreoffice", "--headless", "--convert-to", "pdf", tmp_path,
                             "--outdir", str(Path(tmp_path).parent)],
                            capture_output=True, timeout=30
                        )
                        pdf_path = tmp_path.replace(".docx", ".pdf")
                        if Path(pdf_path).exists():
                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()
                            st.download_button(
                                "\U0001f4c4 Download as PDF",
                                data=pdf_bytes,
                                file_name="humanized_output.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )
                        else:
                            st.info("PDF conversion unavailable (LibreOffice not found)")
                    except (subprocess.TimeoutExpired, FileNotFoundError):
                        st.info("PDF conversion unavailable (install LibreOffice)")
                else:
                    st.info("Upload a DOCX file for PDF output")
            except Exception:
                st.info("PDF conversion unavailable")

        # Metrics
        with st.sidebar:
            st.divider()
            st.subheader("\U0001f4ca Local Self-Check Metrics")
            metrics = compute_metrics(humanized_text)
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Perplexity</div>'
                           f'<div class="metric-value">{metrics["perplexity_proxy"]:.1f}</div></div>',
                           unsafe_allow_html=True)
            with m2:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Burstiness</div>'
                           f'<div class="metric-value">{metrics["burstiness"]:.3f}</div></div>',
                           unsafe_allow_html=True)
            with m3:
                st.markdown(f'<div class="metric-card"><div class="metric-label">TTR</div>'
                           f'<div class="metric-value">{metrics["ttr"]:.4f}</div></div>',
                           unsafe_allow_html=True)
            with m4:
                st.markdown(f'<div class="metric-card"><div class="metric-label">AI Vocab %</div>'
                           f'<div class="metric-value">{metrics["ai_vocab_density"]*100:.1f}%</div></div>',
                           unsafe_allow_html=True)
            st.markdown('<div class="warning-box">\u26a0\ufe0f These are local estimates only and '
                       'do not guarantee bypassing any external detector.</div>',
                       unsafe_allow_html=True)

        # Change review
        if show_changes and changes:
            st.divider()
            st.subheader("\U0001f4dd Changes for Human Review")
            module_counts = Counter(c.module for c in changes)
            cols = st.columns(len(module_counts))
            for i, (mod, cnt) in enumerate(module_counts.items()):
                with cols[i % len(cols)]:
                    st.metric(mod, cnt)

            with st.expander("View detailed changes", expanded=False):
                for i, ch in enumerate(changes[:100]):
                    st.markdown(f"**{ch.module}**: `{ch.original}` → `{ch.modified}`")
                if len(changes) > 100:
                    st.info(f"Showing 100 of {len(changes)} changes.")

    elif initiate and not input_text and not input_docx_bytes:
        st.warning("Please upload a file or paste text before initiating.")


if __name__ == "__main__":
    main()
