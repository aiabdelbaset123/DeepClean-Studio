#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio

A Streamlit app for careful academic text revision. The app focuses on clarity,
flow, citation preservation, transparent change review, and Word export.
"""

from __future__ import annotations

import difflib
import html
import re
from collections import Counter
from dataclasses import dataclass
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

try:
    import textstat
except Exception:  # pragma: no cover - optional quality helper
    textstat = None

try:
    from sentence_transformers import SentenceTransformer, util
except Exception:  # pragma: no cover - optional semantic helper
    SentenceTransformer = None
    util = None


st.set_page_config(page_title="DeepClean Studio", layout="wide")

APP_DIR = Path(__file__).resolve().parent
SYNONYM_COLUMNS = {"domain", "original", "replacement"}


@dataclass(frozen=True)
class RevisionStats:
    words_original: int
    words_revised: int
    sentences_revised: int
    avg_sentence_length: float
    lexical_diversity: float
    similarity: float
    formulaic_phrase_count: int
    certainty_marker_count: int
    sentence_length_variation: float
    authorship_review_band: str


@dataclass(frozen=True)
class SentenceSignal:
    text: str
    score: float
    label: str
    reasons: Tuple[str, ...]


@dataclass(frozen=True)
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


FORMULAIC_PATTERNS = [
    r"\bit is important to note that\b",
    r"\bit should be noted that\b",
    r"\bin today's (?:world|society)\b",
    r"\bplays a (?:crucial|vital) role in\b",
    r"\bhas a significant impact on\b",
    r"\bdelve into\b",
    r"\bstands as\b",
    r"\bserves as\b",
    r"\bis a testament to\b",
    r"\bis a reminder of\b",
    r"\bmarks? a (?:significant|pivotal|key|crucial) (?:moment|shift|turning point)\b",
    r"\b(?:underscores?|highlights?|emphasizes?) (?:the )?(?:importance|significance|need)\b",
    r",\s*(?:highlighting|underscoring|emphasizing|ensuring|reflecting|symbolizing|contributing to|cultivating|fostering|encompassing)\b",
    r"\breflects? broader\b",
    r"\bsymboli[sz]es? (?:its )?(?:ongoing|enduring|lasting)\b",
    r"\bcontribut(?:e|es|ing) to (?:the )?(?:broader|overall)\b",
    r"\bsetting the stage for\b",
    r"\bindelible mark\b",
    r"\bdeeply rooted\b",
    r"\bvaluable insights\b",
    r"\baligns? with\b",
    r"\bresonates? with\b",
    r"\bboasts? a\b",
    r"\brich tapestry\b",
    r"\bevolving landscape\b",
    r"\bactive social media presence\b",
    r"\bindependent coverage\b",
    r"\bmedia outlets\b",
    r"\bprofiled in\b",
    r"\bindustry reports\b",
    r"\bobservers have cited\b",
    r"\bexperts argue\b",
    r"\bsome critics\b",
    r"\bdespite (?:its|these) [^.!?]{0,80}challenges\b",
    r"\bas of my last (?:training update|knowledge update|knowledge)\b",
    r"\bup to my last training update\b",
    r"\bas an ai (?:language model|assistant)\b",
    r"\bi hope this helps\b",
    r"\bof course[!,]?\b",
    r"\bcertainly[!,]?\b",
    r"\bin conclusion\b",
    r"\boverall\b",
    r"\bfurthermore\b",
    r"\bmoreover\b",
    r"\badditionally\b",
    r"\bon the other hand\b",
    r"\bit is worth noting that\b",
    r"\bit should be emphasized that\b",
    r"علاوة على ذلك",
    r"بالإضافة إلى ذلك",
    r"في الختام",
    r"وبناءً على ما سبق",
    r"من ناحية أخرى",
    r"تجدر الإشارة إلى",
]
CERTAINTY_PATTERNS = [
    r"\bclearly proves\b",
    r"\bproves that\b",
    r"\balways\b",
    r"\bnever\b",
    r"\bdefinitely\b",
    r"\bundoubtedly\b",
]

AI_VOCABULARY_TERMS = [
    "additionally",
    "align",
    "aligns with",
    "boasts",
    "comprehensive",
    "contributing",
    "crucial",
    "cultivate",
    "delve",
    "deeply rooted",
    "dynamic",
    "elevate",
    "encompassing",
    "enduring",
    "evolving landscape",
    "foster",
    "furthermore",
    "garner",
    "highlight",
    "holistic",
    "indelible mark",
    "intricate",
    "in conclusion",
    "innovative",
    "landscape",
    "leverage",
    "meticulous",
    "moreover",
    "multifaceted",
    "notably",
    "overall",
    "pivotal",
    "profound",
    "resonate",
    "robust",
    "seamless",
    "serve as",
    "serves as",
    "showcase",
    "significant",
    "stands as",
    "tapestry",
    "testament",
    "transformative",
    "underscore",
    "valuable insights",
    "vibrant",
    "vital",
    "علاوة على ذلك",
    "بالإضافة إلى ذلك",
    "في الختام",
    "بناء على ما سبق",
    "من ناحية أخرى",
    "تجدر الإشارة",
    "بشكل شامل",
    "محوري",
    "حيوي",
]

TRANSPARENCY_COMPONENTS = [
    ("استقبال النص", "تقطيع الكلمات والجمل والتحقق من حد 250 حرفًا."),
    ("عزل النص المؤهل", "استبعاد المراجع والجداول والمعادلات والأكواد قبل حساب مؤشرات الأسلوب."),
    ("الحيرة والتدفق", "قياس تنوع المفردات وتباين أطوال الجمل كمؤشرات محلية."),
    ("محاكاة معيارية محلية", "قواعد قابلة للتفسير تقارن النص بأنماط لغوية شائعة، وليست نموذجًا عميقًا أو خدمة خارجية."),
    ("مصنف الجمل", "إسناد درجة مراجعة لكل جملة بدل الاكتفاء بحكم شامل."),
    ("درع إعادة الصياغة", "رصد اجتماع مفردات آلية مع تدفق منتظم أو صياغة مصقولة جدًا."),
    ("عتبة 20%", "إظهار درع أمان عند الإشارات المنخفضة لتجنب اتهام النصوص البشرية المنظمة."),
    ("ESL De-biasing", "خفض القلق عند ظهور إنجليزية بسيطة ومنظمة دون قوالب آلية واضحة."),
    ("وعي بنية البحث", "تخفيف أثر أقسام المنهجية والمواد لأنها بطبيعتها جافة ومنتظمة."),
    ("الأكواد والترجمة", "إظهار تنبيهات منفصلة عند وجود كتل برمجية أو انتقالات لغوية لا تشبه النثر العادي."),
    ("مفردات الذكاء الاصطناعي", "إبراز الكلمات والعبارات التي ترفع الحاجة إلى مراجعة بشرية."),
]

REALISM_LIMITATIONS = [
    ("لا يوجد اتصال خارجي", "التطبيق لا يتصل بـ GPTZero أو iThenticate أو Turnitin أو Crossref أو PubMed."),
    ("لا حكم قطعي", "النتائج مؤشرات مراجعة فقط، ولا تثبت أن النص بشري أو مولد آليًا."),
    ("لا نسبة أصالة", "النسب المعروضة هي قوة إشارات داخلية، وليست نسبة الجمل المكتوبة بالذكاء الاصطناعي."),
    ("لا تحقق مصادر حقيقي", "تنبيهات المراجع تفحص الشكل محليًا فقط، ولا تؤكد وجود المصدر على الإنترنت."),
    ("قابل للخطأ", "النصوص الأكاديمية الجافة، وكتابة غير الناطقين بالإنجليزية، والترجمات قد ترفع الإشارة خطأً."),
]


def split_sentences(text: str) -> List[str]:
    """Lightweight sentence splitter that avoids runtime model downloads."""
    text = re.sub(r"\s+", " ", text.strip())
    if not text:
        return []
    parts = re.split(r"(?<=[.!?؟])\s+", text)
    return [part.strip() for part in parts if part.strip()]


def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)


def preserve_case(source: str, replacement: str) -> str:
    if source.isupper():
        return replacement.upper()
    if source[:1].isupper():
        return replacement[:1].upper() + replacement[1:]
    return replacement


def normalize_spacing(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"\s+([,.;:!?؟])", r"\1", text)
    text = re.sub(r"([([{])\s+", r"\1", text)
    text = re.sub(r"\s+([])}])", r"\1", text)
    return text.strip()


def strip_chatbot_markup(text: str) -> str:
    """Remove chatbot scaffolding and Markdown wrappers from prose."""
    text = re.sub(r"(?im)^\s*(?:sure|certainly|of course)[!,.\s]+", "", text)
    text = re.sub(
        r"(?im)^\s*here(?:'s| is)\s+(?:a|the)\s+(?:revised|polished|improved|updated)\s+(?:version|draft|text).*?:\s*$",
        "",
        text,
    )
    text = re.sub(r"(?im)^\s*as an ai (?:language model|assistant),?\s*", "", text)
    text = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", text)
    text = re.sub(r"(\*\*|__)([^*_`\n][^*_`\n]*?)\1", r"\2", text)
    text = re.sub(r"(?<!`)`([^`\n]+)`(?!`)", r"\1", text)
    return normalize_spacing(text)


def polish_layout_artifacts(text: str) -> str:
    """Fix small grammar and spacing artifacts created by rule-based revision."""
    text = normalize_spacing(text)
    text = re.sub(r"\b(an)\s+([bcdfghjklmnpqrstvwxyz][A-Za-z-]*)", r"a \2", text, flags=re.I)
    text = re.sub(r"\b(a)\s+([aeio][A-Za-z-]*)", r"an \2", text, flags=re.I)
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b(an|a)\s+(?:social media accounts)\b", "social media accounts", text, flags=re.I)
    text = re.sub(r"\s+([,.;:!?؟،])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = re.sub(r"([.!?])\s+([a-z])", lambda m: f"{m.group(1)} {m.group(2).upper()}", text)
    text = re.sub(r"([؟])\s+([a-zA-Z])", lambda m: f"{m.group(1)} {m.group(2).upper()}", text)
    return normalize_spacing(text)


@st.cache_data
def load_synonym_dictionary(filepath: str = "synonyms_academic.csv") -> pd.DataFrame:
    path = APP_DIR / filepath
    if not path.exists():
        st.warning("ملف المرادفات غير موجود؛ سيعمل التطبيق بدون استبدالات مجال تخصصية.")
        return pd.DataFrame(columns=sorted(SYNONYM_COLUMNS))

    try:
        df = pd.read_csv(path, encoding="utf-8")
    except UnicodeDecodeError:
        df = pd.read_csv(path, encoding="utf-8-sig")
    except Exception as exc:
        st.error(f"تعذر تحميل ملف المرادفات: {exc}")
        return pd.DataFrame(columns=sorted(SYNONYM_COLUMNS))

    missing = SYNONYM_COLUMNS.difference(df.columns)
    if missing:
        st.error("ملف المرادفات يجب أن يحتوي على الأعمدة: domain, original, replacement")
        return pd.DataFrame(columns=sorted(SYNONYM_COLUMNS))

    df = df.dropna(subset=["domain", "original", "replacement"]).copy()
    for column in SYNONYM_COLUMNS:
        df[column] = df[column].astype(str).str.strip()
    return df[df["original"].ne("") & df["replacement"].ne("")]


@st.cache_resource
def load_semantic_model():
    if SentenceTransformer is None:
        return None
    try:
        return SentenceTransformer("all-MiniLM-L6-v2")
    except Exception:
        st.warning("تعذر تحميل نموذج القفل الدلالي؛ سيستخدم التطبيق فحص تشابه محليًا أخف.")
        return None


class AcademicRevisionEngine:
    def __init__(
        self,
        domain: str,
        intensity: int,
        text: str,
        preserve_word_count: bool,
    ) -> None:
        self.domain = domain
        self.intensity = intensity
        self.original_text = text
        self.preserve_word_count = preserve_word_count
        self.synonym_map = self._build_synonym_map()
        self.semantic_model = load_semantic_model()
        self.citation_pattern = re.compile(r"\[[\d,\-; ]+\]|\([^)]*\d{4}[^)]*\)")
        self.data_pattern = re.compile(
            r"\b\d+(?:\.\d+)?\s?(?:%|mg|g|kg|ml|L|mm|cm|m|km|Hz|kHz|MHz|GHz|V|W|kW|°C|K|s|min|h)?\b",
            re.I,
        )
        self.reference_lengths = [15, 22, 18, 30, 25, 19, 28, 16, 21, 27, 23, 17, 20, 24, 26, 19, 22, 31]
        self.hedges = {
            "medical": ["may indicate", "appears to suggest", "is clinically consistent with"],
            "engineering": ["may indicate", "appears to support", "is consistent with"],
            "humanities": ["may suggest", "can be read as", "appears to reflect"],
            "general": ["may suggest", "appears to indicate", "is consistent with"],
        }
        self.causal_links = {
            "medical": {"therefore": "as a result", "so": "as a result", "because": "because of this"},
            "engineering": {"therefore": "consequently", "so": "as a result", "because": "due to this"},
            "humanities": {"therefore": "consequently", "so": "in this context", "because": "owing to this"},
            "general": {"therefore": "accordingly", "so": "as a result", "because": "because of this"},
        }
        self.contrast_links = {
            "medical": {"but": "whereas", "however": "in contrast"},
            "engineering": {"but": "whereas", "however": "alternatively"},
            "humanities": {"but": "nevertheless", "however": "nonetheless"},
            "general": {"but": "however", "however": "nevertheless"},
        }
        self.openers = {
            "medical": ["Clinically", "In this finding", "From a diagnostic perspective"],
            "engineering": ["In practice", "At the system level", "From a design perspective"],
            "humanities": ["In this context", "Historically", "At the interpretive level"],
            "general": ["In this context", "More specifically", "At this stage"],
        }
        self.formulaic_rewrites = [
            (r"\bclearly proves that\b", "suggests that"),
            (r"\bproves that\b", "suggests that"),
            (r"\bclearly demonstrates that\b", "indicates that"),
            (r"\bit is important to note that\s+", ""),
            (r"\bit should be noted that\s+", ""),
            (r"\bit is worth noting that\s+", ""),
            (r"\bit should be emphasized that\s+", ""),
            (r"\bin today's (?:world|society)\b", "currently"),
            (r"\bplays a (?:crucial|vital) role in\b", "is part of"),
            (r"\bhas a significant impact on\b", "affects"),
            (r"\bdelve into\b", "examine"),
            (r"\bstands as a testament to the evolving landscape of\b", "is part of"),
            (r"\bserves as a testament to the evolving landscape of\b", "is part of"),
            (r"\bstands as a testament to\b", "shows"),
            (r"\bserves as a testament to\b", "shows"),
            (r"\bmaintains an active social media presence\b", "has social media accounts"),
            (r"\bstands as (?:a|an|the)?\s*", "is "),
            (r"\bserves as (?:a|an|the)?\s*", "is "),
            (r"\bis testament to\b", "shows"),
            (r"\bis a testament to\b", "shows"),
            (r"\bis a reminder of\b", "shows"),
            (r",\s*highlighting its (?:pivotal|key|crucial|vital) role and significant impact on the broader ([^.!?]+)", r" and affects the \1"),
            (r",\s*(?:highlighting|underscoring|emphasizing) (?:the )?(?:importance|significance) of ([^.!?]+)", r" and discusses \1"),
            (r"\bpivotal role\b", "role"),
            (r"\bsignificant impact\b", "effect"),
            (r"\bbroader community\b", "community"),
            (r"\bmarks? a (?:significant|pivotal|key|crucial) (?:moment|shift|turning point) in\b", "changed"),
            (r"\b(?:underscores?|highlights?|emphasizes?) (?:the )?(?:importance|significance|need) of\b", "shows"),
            (r"\breflects? broader\b", "reflects"),
            (r"\bsymboli[sz]es? (?:its )?(?:ongoing|enduring|lasting)\b", "shows"),
            (r"\bsetting the stage for\b", "preceding"),
            (r"\bindelible mark\b", "effect"),
            (r"\bdeeply rooted\b", "long-standing"),
            (r"\bvaluable insights into\b", "evidence about"),
            (r"\baligns? with\b", "matches"),
            (r"\bresonates? with\b", "relates to"),
            (r"\bboasts? a\b", "has a"),
            (r"\brich tapestry of\b", "range of"),
            (r"\bevolving landscape\b", "field"),
            (r"\bactive social media presence\b", "social media accounts"),
            (r"\bindependent coverage\b", "coverage"),
            (r"\bmedia outlets\b", "publications"),
            (r"\bprofiled in\b", "covered in"),
            (r"\bindustry reports\b", "reports"),
            (r"\bobservers have cited\b", "sources cite"),
            (r"\bexperts argue\b", "some sources argue"),
            (r"\bsome critics\b", "some reviewers"),
            (r"\bas of my last (?:training update|knowledge update|knowledge),?\s*", ""),
            (r"\bup to my last training update,?\s*", ""),
            (r"\bas an ai (?:language model|assistant),?\s*", ""),
            (r"\bi hope this helps[.!]?\s*", ""),
            (r"^\s*(?:of course|certainly)[!,]?\s*", ""),
            (r"\butilize\b", "use"),
            (r"\bfacilitate\b", "support"),
            (r"\bmoreover,\s+moreover,\s+", "moreover, "),
            (r"\bin conclusion,\s+", ""),
            (r"\boverall,\s+", ""),
        ]
        self.repetitive_openers = {
            "furthermore": "More specifically",
            "moreover": "In addition",
            "additionally": "At the same time",
            "therefore": "Accordingly",
            "however": "In contrast",
        }
        self.editorial_connectors = [
            r"\bfurthermore,?\s+",
            r"\bmoreover,?\s+",
            r"\badditionally,?\s+",
            r"\bin conclusion,?\s+",
            r"\boverall,?\s+",
            r"\bon the other hand,?\s+",
            r"\bit is worth noting that\s+",
            r"\bit should be emphasized that\s+",
            r"\bas previously mentioned,?\s+",
            r"\bbased on the foregoing,?\s+",
        ]
        self.editorial_vocabulary_swaps = [
            (r"\bshows important results\b", "reports specific findings"),
            (r"\bshows that\b", "indicates that"),
            (r"\bimportant\b", "specific"),
            (r"\bsignificant\b", "measurable"),
            (r"\bshows\b", "reveals"),
            (r"\bindicates\b", "suggests"),
            (r"\bresults\b", "findings"),
            (r"\bdata\b", "evidence"),
            (r"\bmethod\b", "approach"),
            (r"\btopic\b", "question"),
            (r"\bfactor\b", "force"),
            (r"\ba large number of\b", "a broad range of"),
            (r"\bin a clear way\b", "with unusual clarity"),
        ]
        self.editorial_passive_rewrites = [
            (r"\bit was found that\b", "I read the evidence as suggesting that"),
            (r"\bit was observed that\b", "the evidence brings into view that"),
            (r"\bit is argued that\b", "I would argue that"),
            (r"\bit is suggested that\b", "the material suggests that"),
            (r"\bit can be seen that\b", "we can see that"),
            (r"\bwas analyzed\b", "comes under close examination"),
            (r"\bwere analyzed\b", "come under close examination"),
            (r"\bwas examined\b", "receives close attention"),
            (r"\bwere examined\b", "receive close attention"),
        ]
        self.editorial_voice_openers = [
            "In this passage",
            "In the evidence",
            "In the analysis",
            "In the reported material",
        ]
        self.editorial_short_beats: List[str] = []
        self.editorial_question_added = False

    def _build_synonym_map(self) -> Dict[str, str]:
        df = load_synonym_dictionary()
        if df.empty:
            return {}
        if self.domain == "general":
            scoped = df
        else:
            scoped = df[df["domain"].str.lower() == self.domain.lower()]
        return {
            str(row["original"]).lower(): str(row["replacement"])
            for _, row in scoped.iterrows()
            if str(row["original"]).strip() and str(row["replacement"]).strip()
        }

    def _protect_fragments(self, text: str) -> Tuple[str, Dict[str, str]]:
        replacements: Dict[str, str] = {}

        def replace(match: re.Match[str]) -> str:
            token = f"__PROTECTED_{len(replacements)}__"
            replacements[token] = match.group(0)
            return token

        protected = self.citation_pattern.sub(replace, text)
        protected = self.data_pattern.sub(replace, protected)
        return protected, replacements

    def _restore_fragments(self, text: str, replacements: Dict[str, str]) -> str:
        for token, value in replacements.items():
            text = text.replace(token, value)
        return text

    def _semantic_similarity(self, original: str, revised: str) -> float:
        if not original.strip() or not revised.strip():
            return 0.0
        if self.semantic_model is not None and util is not None:
            emb_original = self.semantic_model.encode(original, convert_to_tensor=True)
            emb_revised = self.semantic_model.encode(revised, convert_to_tensor=True)
            return float(util.pytorch_cos_sim(emb_original, emb_revised).item())
        original_words = [word.lower() for word in tokenize_words(original)]
        revised_words = [word.lower() for word in tokenize_words(revised)]
        jaccard = jaccard_similarity(original_words, revised_words)
        sequence = difflib.SequenceMatcher(None, original.lower(), revised.lower()).ratio()
        return max(jaccard, sequence)

    def _semantic_lock(self, original: str, candidate: str, threshold: float = 0.92) -> str:
        effective_threshold = threshold if self.semantic_model is not None else min(threshold, 0.35)
        if self._semantic_similarity(original, candidate) >= effective_threshold:
            return candidate
        return original

    def _most_predictable_terms(self, sentence: str) -> Sequence[str]:
        words = [word for word in tokenize_words(sentence) if word.lower() in self.synonym_map]
        if not words:
            return []
        counts = Counter(word.lower() for word in tokenize_words(self.original_text))
        return sorted(words, key=lambda word: (-counts[word.lower()], -len(word)))

    def _rewrite_formulaic_phrases(self, sentence: str) -> str:
        revised = sentence
        for pattern, replacement in self.formulaic_rewrites:
            revised = re.sub(pattern, replacement, revised, flags=re.I)
        revised = normalize_spacing(revised)
        return revised[:1].upper() + revised[1:] if revised else revised

    def _revise_repetitive_opener(self, sentence: str, index: int) -> str:
        if index == 0:
            return sentence
        for source, replacement in self.repetitive_openers.items():
            if re.match(rf"^\s*{re.escape(source)},?\s+", sentence, flags=re.I):
                if index % 2 == 0:
                    return re.sub(rf"^\s*{re.escape(source)},?\s+", f"{replacement}, ", sentence, count=1, flags=re.I)
                shortened = re.sub(rf"^\s*{re.escape(source)},?\s+", "", sentence, count=1, flags=re.I)
                return shortened[:1].upper() + shortened[1:]
        return sentence

    def engine1_perplexity_injector(self, sentence: str) -> str:
        """Conservative term revision and hedging for over-certain academic claims."""
        revised = self._rewrite_formulaic_phrases(sentence)
        if not self.synonym_map or self.intensity < 2:
            revised = revised
        else:
            max_changes = max(1, min(3, self.intensity - 1))
            for word in self._most_predictable_terms(revised)[:max_changes]:
                replacement = self.synonym_map.get(word.lower())
                if replacement:
                    revised = re.sub(
                        rf"\b{re.escape(word)}\b",
                        preserve_case(word, replacement),
                        revised,
                        count=1,
                        flags=re.I,
                    )

        certainty_markers = re.compile(r"\b(always|never|definitely|undoubtedly)\b", re.I)
        if self.intensity >= 3 and certainty_markers.search(revised):
            hedge = self.hedges.get(self.domain, self.hedges["general"])[0]
            revised = certainty_markers.sub(hedge, revised, count=1)
        elif textstat is not None and self.intensity >= 4:
            try:
                if textstat.flesch_reading_ease(revised) < 20 and "," not in revised:
                    hedge = self.hedges.get(self.domain, self.hedges["general"])[1]
                    revised = f"{hedge}, {revised[:1].lower()}{revised[1:]}"
            except Exception:
                pass
        return normalize_spacing(revised)

    def engine2_burstiness_synthesizer(self, sentences: List[str]) -> List[str]:
        """Adjust overly long sentences and avoid repeated adjacent sentence lengths."""
        adjusted: List[str] = []
        for sentence in sentences:
            adjusted.extend(self._split_long_sentence(sentence))

        balanced: List[str] = []
        previous_len: Optional[int] = None
        for sentence in adjusted:
            words = sentence.split()
            current_len = len(tokenize_words(sentence))
            if previous_len is not None and abs(current_len - previous_len) <= 2 and current_len > 18:
                midpoint = max(10, len(words) // 2)
                first = " ".join(words[:midpoint]).rstrip(" ,;")
                second = " ".join(words[midpoint:]).strip(" ,;")
                if first and second:
                    if first[-1] not in ".!?؟":
                        first += "."
                    second = second[:1].upper() + second[1:]
                    if second[-1] not in ".!?؟":
                        second += "."
                    balanced.extend([first, second])
                    previous_len = len(tokenize_words(second))
                    continue
            balanced.append(sentence)
            previous_len = current_len
        return balanced

    def engine3_style_variety_editor(self, sentences: List[str]) -> List[str]:
        """Vary punctuation and sentence openings without adding new information."""
        if self.intensity < 3:
            return [self._revise_repetitive_opener(sentence, index) for index, sentence in enumerate(sentences)]
        openers = self.openers.get(self.domain, self.openers["general"])
        revised: List[str] = []
        for index, sentence in enumerate(sentences):
            sentence = self._revise_repetitive_opener(sentence, index)
            if index > 0 and index % 4 == 0 and len(tokenize_words(sentence)) > 9:
                opener = openers[index % len(openers)]
                if not sentence.lower().startswith(opener.lower()):
                    sentence = f"{opener}, {sentence[:1].lower()}{sentence[1:]}"
            if self.intensity >= 4 and ";" not in sentence and "," in sentence and index % 5 == 2:
                sentence = sentence.replace(",", ";", 1)
            revised.append(sentence)
        return revised

    def engine4_semantic_deepener(self, sentences: List[str]) -> List[str]:
        """Replace generic connectors with more precise causal or contrastive links."""
        causal = self.causal_links.get(self.domain, self.causal_links["general"])
        contrast = self.contrast_links.get(self.domain, self.contrast_links["general"])
        revised: List[str] = []
        for sentence in sentences:
            original = sentence
            for source, target in {**causal, **contrast}.items():
                sentence = re.sub(rf"^\s*{re.escape(source)}\b", target, sentence, flags=re.I)
                sentence = re.sub(rf"\b{re.escape(source)}\b", target, sentence, count=1, flags=re.I)
            revised.append(self._semantic_lock(original, sentence, threshold=0.75))
        return revised

    def engine5_structure_regularizer(self, sentences: List[str]) -> List[str]:
        """Reduce repetitive paragraph rhythm while preserving sentence order."""
        revised: List[str] = []
        previous_start = ""
        for sentence in sentences:
            cleaned = re.sub(r"\b(very|really|basically|clearly|simply|obviously)\b", "", sentence, flags=re.I)
            words = tokenize_words(cleaned)
            current_start = " ".join(word.lower() for word in words[:3])
            if previous_start and current_start == previous_start and len(words) > 6:
                cleaned = " ".join(words[3:])
                cleaned = cleaned[:1].upper() + cleaned[1:]
            previous_start = current_start
            revised.append(normalize_spacing(cleaned))
        return revised

    def engine6_coherence_checker(self, original_sentences: List[str], revised_sentences: List[str]) -> List[str]:
        """Repair obvious formatting, protected-token, and sentence-ending issues."""
        checked: List[str] = []
        protected_token = re.compile(r"__PROTECTED_\d+__")
        for index, sentence in enumerate(revised_sentences):
            original = original_sentences[min(index, len(original_sentences) - 1)] if original_sentences else sentence
            if protected_token.findall(original) != protected_token.findall(sentence):
                sentence = original
            sentence = normalize_spacing(sentence)
            if sentence and sentence[-1] not in ".!?؟":
                sentence += "."
            checked.append(self._semantic_lock(original, sentence, threshold=0.55))
        return checked

    def _split_long_sentence(self, sentence: str) -> List[str]:
        words = tokenize_words(sentence)
        limit = 42 if self.intensity <= 2 else 34
        if len(words) <= limit:
            return [sentence]

        breakpoints = [",", ";", " and ", " but ", " whereas ", " which "]
        lower = sentence.lower()
        midpoint = len(sentence) // 2
        candidates: List[int] = []
        for marker in breakpoints:
            start = 0
            while True:
                idx = lower.find(marker, start)
                if idx == -1:
                    break
                candidates.append(idx + len(marker.strip()))
                start = idx + 1

        if candidates:
            split_at = min(candidates, key=lambda idx: abs(idx - midpoint))
            first = sentence[:split_at].strip(" ,;")
            second = sentence[split_at:].strip(" ,;")
        else:
            raw_words = sentence.split()
            split_at_word = max(12, len(raw_words) // 2)
            first = " ".join(raw_words[:split_at_word]).strip(" ,;")
            second = " ".join(raw_words[split_at_word:]).strip(" ,;")

        if not first or not second:
            return [sentence]
        if first[-1] not in ".!?؟":
            first += "."
        second = second[:1].upper() + second[1:]
        if second[-1] not in ".!?؟":
            second += "."
        return [first, second]

    def _flow_is_extreme(self, sentences: List[str]) -> bool:
        lengths = [len(tokenize_words(sentence)) for sentence in sentences]
        if len(lengths) < 3:
            return False
        baseline = np.std(self.reference_lengths)
        return bool(np.std(lengths) > 3 * baseline)

    def _clean_editorial_connectors(self, sentence: str) -> str:
        sentence = sentence.replace("—", ",").replace("–", ",")
        for pattern in self.editorial_connectors:
            sentence = re.sub(pattern, "", sentence, flags=re.I)
        sentence = re.sub(r"\s*[-]{2,}\s*", ", ", sentence)
        return normalize_spacing(sentence)

    def _activate_english_voice(self, sentence: str) -> str:
        for pattern, replacement in self.editorial_passive_rewrites:
            sentence = re.sub(pattern, replacement, sentence, flags=re.I)
        return normalize_spacing(sentence)

    def _enrich_english_vocabulary(self, sentence: str) -> str:
        limit = 1 if self.intensity < 4 else 2
        changes = 0
        for pattern, replacement in self.editorial_vocabulary_swaps:
            if changes >= limit:
                break
            if re.search(pattern, sentence, flags=re.I):
                sentence = re.sub(pattern, replacement, sentence, count=1, flags=re.I)
                changes += 1
        return normalize_spacing(sentence)

    def _polish_english_artifacts(self, sentence: str) -> str:
        sentence = re.sub(r"\breveals revealing (?:results|findings)\b", "reveals meaningful findings", sentence, flags=re.I)
        sentence = re.sub(r"\b(clinical measurements|evidence|data) reveals\b", r"\1 reveal", sentence, flags=re.I)
        sentence = re.sub(r"\bI read the evidence as suggesting that the evidence\b", "I read the evidence as suggesting that it", sentence, flags=re.I)
        return normalize_spacing(sentence)

    def _english_editorial_pass(self, paragraph: str) -> str:
        if self.intensity < 3:
            return normalize_spacing(paragraph.replace("—", ",").replace("–", ","))

        sentences = split_sentences(paragraph)
        if not sentences:
            return paragraph

        revised: List[str] = []
        for index, sentence in enumerate(sentences):
            sentence = self._clean_editorial_connectors(sentence)
            sentence = self._activate_english_voice(sentence)
            sentence = self._enrich_english_vocabulary(sentence)
            sentence = self._polish_english_artifacts(sentence)
            sentence = sentence[:1].upper() + sentence[1:] if sentence else sentence
            if sentence and sentence[-1] not in ".!?":
                sentence += "."
            revised.append(sentence)

            if self.editorial_short_beats and self.intensity >= 5 and len(tokenize_words(sentence)) >= 20 and index % 3 == 0:
                revised.append(self.editorial_short_beats[index % len(self.editorial_short_beats)])

        return normalize_spacing(" ".join(revised))

    def revise_paragraph(self, paragraph: str, context: Sequence[str] | None = None) -> str:
        sentences = split_sentences(paragraph)
        context = list(context or [])
        working = [self.engine1_perplexity_injector(sentence) for sentence in sentences]
        working = self.engine2_burstiness_synthesizer(working)
        working = self.engine3_style_variety_editor(working)
        working = self.engine4_semantic_deepener(working)
        working = self.engine5_structure_regularizer(working)
        working = self.engine6_coherence_checker(sentences, working)

        if self._flow_is_extreme(working):
            working = self.engine2_burstiness_synthesizer(sentences)
            working = self.engine6_coherence_checker(sentences, working)

        paragraph = normalize_spacing(" ".join(working))
        if self.semantic_model is not None:
            paragraph = self._semantic_lock(" ".join(context + sentences), " ".join(context) + " " + paragraph, threshold=0.70)
            if context and paragraph.startswith(" ".join(context)):
                paragraph = paragraph[len(" ".join(context)):].strip()
        paragraph = self._english_editorial_pass(paragraph)
        if self.preserve_word_count:
            paragraph = self._avoid_expansion(paragraph)
        return paragraph

    def _avoid_expansion(self, text: str) -> str:
        original_count = len(tokenize_words(self.original_text))
        revised_count = len(tokenize_words(text))
        if not original_count or revised_count <= original_count * 1.15:
            return text
        text = re.sub(
            r"\b(Furthermore|Additionally|Moreover|However|Therefore|Notably|Clinically|Historically|In this passage|In the evidence|In the analysis|In the reported material),\s+",
            "",
            text,
        )
        return normalize_spacing(text)

    def run(self) -> str:
        text = strip_chatbot_markup(self.original_text)
        protected, fragments = self._protect_fragments(text)
        paragraphs = [part.strip() for part in re.split(r"\n\s*\n", protected) if part.strip()]
        revised = []
        context: List[str] = []
        for paragraph in paragraphs:
            revised_paragraph = self.revise_paragraph(paragraph, context=context)
            revised.append(revised_paragraph)
            context = split_sentences(paragraph)[-5:]
        final_text = "\n\n".join(revised)
        final_text = self._restore_fragments(final_text, fragments)
        return normalize_spacing(final_text)


def is_mostly_arabic(text: str) -> bool:
    letters = re.findall(r"[A-Za-z\u0600-\u06FF]", text)
    if not letters:
        return False
    arabic_letters = re.findall(r"[\u0600-\u06FF]", text)
    return len(arabic_letters) / len(letters) >= 0.35


class ArabicEditorialRevisionEngine:
    """Rule-based Arabic prose editor for warmer, less mechanical rewriting."""

    mechanical_connectors = [
        r"علاوة على ذلك[،,]?\s*",
        r"بالإضافة إلى ذلك[،,]?\s*",
        r"في الختام[،,]?\s*",
        r"وبناءً على ما سبق[،,]?\s*",
        r"بناءً على ما سبق[،,]?\s*",
        r"من ناحية أخرى[،,]?\s*",
        r"تجدر الإشارة إلى أن?\s*",
        r"ومن الجدير بالذكر أن?\s*",
        r"كما يجب التنويه إلى أن?\s*",
    ]
    vocabulary_swaps = [
        (r"\bبشكل واضح\b", "بجلاء"),
        (r"\bمهم\b", "لافت"),
        (r"\bمهمة\b", "لافتة"),
        (r"\bجيد\b", "متين"),
        (r"\bكبير\b", "واسع الأثر"),
        (r"\bواضح\b", "جلي"),
        (r"\bيؤثر على\b", "يمس"),
        (r"\bيساهم في\b", "يغذي"),
        (r"\bيعكس\b", "يكشف"),
        (r"\bيوضح\b", "يجلي"),
        (r"\bيظهر\b", "يطفو"),
        (r"\bالموضوع\b", "المسألة"),
        (r"\bالنتائج\b", "الحصيلة"),
        (r"\bالبيانات\b", "المعطيات"),
        (r"\bالطريقة\b", "النهج"),
        (r"\bالفكرة\b", "اللمحة"),
        (r"\bالعامل\b", "المؤثر"),
        (r"\bعدد كبير من\b", "طيف واسع من"),
        (r"\bبشكل عام\b", "في المشهد الأوسع"),
    ]
    passive_rewrites = [
        (r"\bتمت دراسة\b", "تعاين الدراسة"),
        (r"\bتمت ملاحظة\b", "تلمح المعاينة"),
        (r"\bتمت مناقشة\b", "يناقش النص"),
        (r"\bتم تحليل\b", "نعاين"),
        (r"\bتم استخدام\b", "يوظف النص"),
        (r"\bتم الاعتماد على\b", "يتكئ النص على"),
        (r"\bتم التركيز على\b", "نقترب من"),
    ]
    voice_openers = [
        "في تقديري",
        "من واقع المعاينة",
        "ما يبدو لي جلياً",
        "حين أقرأ المشهد عن قرب",
    ]
    short_beats = [
        "هنا يتغير الإيقاع.",
        "وهنا بيت القصيد.",
        "الأمر ليس عابراً.",
        "هذه ليست زينة.",
    ]

    def __init__(self, intensity: int, text: str, preserve_word_count: bool) -> None:
        self.intensity = intensity
        self.original_text = text
        self.preserve_word_count = preserve_word_count
        self.question_added = False

    def _split_arabic_sentences(self, text: str) -> List[str]:
        text = re.sub(r"\s+", " ", text.strip())
        if not text:
            return []
        parts = re.split(r"(?<=[.!?؟])\s+", text)
        return [part.strip() for part in parts if part.strip()]

    def _clean_mechanical_transitions(self, sentence: str) -> str:
        sentence = sentence.replace("—", "،").replace("–", "،")
        for pattern in self.mechanical_connectors:
            sentence = re.sub(pattern, "", sentence, flags=re.I)
        sentence = re.sub(r"\s*[-]{2,}\s*", "، ", sentence)
        return normalize_spacing(sentence)

    def _enrich_vocabulary(self, sentence: str, index: int) -> str:
        limit = 1 if self.intensity < 4 else 2
        changes = 0
        for pattern, replacement in self.vocabulary_swaps:
            if changes >= limit:
                break
            if re.search(pattern, sentence):
                sentence = re.sub(pattern, replacement, sentence, count=1)
                changes += 1
        return normalize_spacing(sentence)

    def _activate_voice(self, sentence: str) -> str:
        for pattern, replacement in self.passive_rewrites:
            sentence = re.sub(pattern, replacement, sentence)
        return normalize_spacing(sentence)

    def _shape_rhythm(self, sentences: List[str]) -> List[str]:
        if len(sentences) < 2:
            return sentences
        shaped: List[str] = []
        for index, sentence in enumerate(sentences):
            shaped.append(sentence)
            words_count = len(tokenize_words(sentence))
            should_add_beat = bool(self.short_beats) and self.intensity >= 4 and words_count >= 18 and index % 3 == 0
            if should_add_beat:
                shaped.append(self.short_beats[index % len(self.short_beats)])
        return shaped

    def _add_mid_text_question(self, sentences: List[str]) -> List[str]:
        return sentences

    def _avoid_expansion(self, text: str) -> str:
        if not self.preserve_word_count:
            return text
        original_count = len(tokenize_words(self.original_text))
        revised_count = len(tokenize_words(text))
        if not original_count or revised_count <= original_count * 1.25:
            return text
        text = re.sub(r"\b(?:في تقديري|من واقع المعاينة|ما يبدو لي جلياً)،\s*", "", text, count=1)
        text = re.sub(r"(?:هنا يتغير الإيقاع|وهنا بيت القصيد|الأمر ليس عابراً|هذه ليست زينة)\.", "", text)
        return normalize_spacing(text)

    def revise_paragraph(self, paragraph: str) -> str:
        sentences = self._split_arabic_sentences(paragraph)
        revised: List[str] = []
        for index, sentence in enumerate(sentences):
            sentence = self._clean_mechanical_transitions(sentence)
            sentence = self._activate_voice(sentence)
            sentence = self._enrich_vocabulary(sentence, index)
            if sentence and sentence[-1] not in ".!?؟":
                sentence += "."
            revised.append(sentence)
        revised = self._shape_rhythm(revised)
        revised = self._add_mid_text_question(revised)
        return normalize_spacing(" ".join(revised))

    def run(self) -> str:
        text = strip_chatbot_markup(self.original_text)
        paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text) if part.strip()]
        revised = [self.revise_paragraph(paragraph) for paragraph in paragraphs]
        final_text = "\n\n".join(revised)
        final_text = final_text.replace("—", "،").replace("–", "،")
        return self._avoid_expansion(normalize_spacing(final_text))


SCIENTIFIC_SECTION_NAMES = {
    "abstract",
    "keywords",
    "introduction",
    "background",
    "literature review",
    "materials and methods",
    "methods",
    "methodology",
    "results",
    "discussion",
    "conclusion",
    "conclusions",
    "acknowledgements",
    "acknowledgments",
    "references",
    "الملخص",
    "الكلمات المفتاحية",
    "المقدمة",
    "الخلفية",
    "الدراسات السابقة",
    "المواد والطرق",
    "المنهجية",
    "النتائج",
    "المناقشة",
    "الخاتمة",
    "الاستنتاجات",
    "المراجع",
}


def set_word_font(run, font_name: str = "Times New Roman", size: int = 12, bold: bool = False) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:cs"), "Arial")


def set_word_bidi(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)
    bidi.set(qn("w:val"), "1")


def mostly_rtl_text(text: str) -> bool:
    letters = re.findall(r"[A-Za-z\u0600-\u06FF]", text)
    if not letters:
        return False
    return len(re.findall(r"[\u0600-\u06FF]", text)) / len(letters) >= 0.35


def academic_line_kind(line: str) -> str:
    stripped = line.strip()
    if not stripped:
        return "blank"
    if stripped.count("|") >= 2 or "\t" in stripped:
        return "table"
    if re.match(r"^\s*(?:[-*•]|\d+[.)])\s+\S+", stripped):
        return "list"
    cleaned = re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", stripped).strip(":：.- ").lower()
    words = tokenize_words(stripped)
    terminal = stripped[-1:] in ".!?؟،؛"
    if cleaned in SCIENTIFIC_SECTION_NAMES:
        return "heading"
    if 1 <= len(words) <= 12 and not terminal and not re.search(r",|;|،|؛", stripped):
        return "heading"
    return "paragraph"


def setup_academic_document(doc: Document) -> None:
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


def add_academic_paragraph(doc: Document, line: str, kind: str) -> None:
    rtl = mostly_rtl_text(line)
    if kind == "heading":
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(12)
        paragraph.paragraph_format.space_after = Pt(6)
        if rtl:
            set_word_bidi(paragraph)
        run = paragraph.add_run(line.strip())
        set_word_font(run, size=13, bold=True)
        return

    if kind == "list":
        marker = re.match(r"^\s*(?:[-*•]|\d+[.)])\s+", line)
        content = line[marker.end() :].strip() if marker else line.strip()
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.LEFT
        if rtl:
            set_word_bidi(paragraph)
        run = paragraph.add_run(content)
        set_word_font(run)
        return

    if kind == "table":
        cells = [cell.strip() for cell in re.split(r"\||\t", line) if cell.strip()]
        table = doc.add_table(rows=1, cols=max(1, len(cells)))
        table.style = "Table Grid"
        for idx, value in enumerate(cells):
            cell = table.cell(0, idx)
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if mostly_rtl_text(value) else WD_ALIGN_PARAGRAPH.LEFT
            if mostly_rtl_text(value):
                set_word_bidi(p)
            run = p.add_run(value)
            set_word_font(run, size=10)
        doc.add_paragraph()
        return

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if rtl else WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = Inches(0 if rtl else 0.25)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.15
    if rtl:
        set_word_bidi(paragraph)
    run = paragraph.add_run(line.strip())
    set_word_font(run)


def create_word_document(text: str, title: Optional[str] = None) -> BytesIO:
    doc = Document()
    setup_academic_document(doc)
    if title:
        heading = doc.add_paragraph()
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = heading.add_run(title.strip())
        set_word_font(run, size=14, bold=True)
    previous_blank = False
    for raw_line in polish_layout_artifacts(text).split("\n"):
        kind = academic_line_kind(raw_line)
        if kind == "blank":
            if not previous_blank:
                doc.add_paragraph()
            previous_blank = True
            continue
        add_academic_paragraph(doc, raw_line, kind)
        previous_blank = False
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


def render_academic_preview(text: str) -> str:
    parts: List[str] = ["<div class='academic-preview'>"]
    in_list = False
    for raw_line in polish_layout_artifacts(text).split("\n"):
        line = raw_line.strip()
        kind = academic_line_kind(raw_line)
        if kind != "list" and in_list:
            parts.append("</ul>")
            in_list = False
        if kind == "blank":
            continue
        if kind == "heading":
            parts.append(f"<h3>{html.escape(line)}</h3>")
        elif kind == "list":
            marker = re.match(r"^\s*(?:[-*•]|\d+[.)])\s+", line)
            content = line[marker.end() :].strip() if marker else line
            if not in_list:
                parts.append("<ul>")
                in_list = True
            parts.append(f"<li>{html.escape(content)}</li>")
        elif kind == "table":
            cells = [html.escape(cell.strip()) for cell in re.split(r"\||\t", line) if cell.strip()]
            parts.append("<table><tr>" + "".join(f"<td>{cell}</td>" for cell in cells) + "</tr></table>")
        else:
            rtl_attr = " dir='rtl'" if mostly_rtl_text(line) else ""
            parts.append(f"<p{rtl_attr}>{html.escape(line)}</p>")
    if in_list:
        parts.append("</ul>")
    parts.append("</div>")
    return "".join(parts)


def word_level_diff(original: str, modified: str) -> str:
    orig_words = original.split()
    mod_words = modified.split()
    matcher = difflib.SequenceMatcher(None, orig_words, mod_words)
    html_parts: List[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        old = html.escape(" ".join(orig_words[i1:i2]))
        new = html.escape(" ".join(mod_words[j1:j2]))
        if tag == "equal":
            html_parts.append(old)
        elif tag == "replace":
            html_parts.append(f"<span class='removed'>{old}</span> <span class='added'>{new}</span>")
        elif tag == "delete":
            html_parts.append(f"<span class='removed'>{old}</span>")
        elif tag == "insert":
            html_parts.append(f"<span class='added'>{new}</span>")
    return " ".join(part for part in html_parts if part)


def jaccard_similarity(original: Iterable[str], revised: Iterable[str]) -> float:
    original_set = set(original)
    revised_set = set(revised)
    if not original_set and not revised_set:
        return 1.0
    return len(original_set & revised_set) / max(1, len(original_set | revised_set))


def count_pattern_matches(text: str, patterns: Sequence[str]) -> int:
    return sum(len(re.findall(pattern, text, flags=re.I)) for pattern in patterns)


def authorship_review_band(formulaic_count: int, certainty_count: int, sentence_variation: float) -> str:
    signals = 0
    if formulaic_count:
        signals += 1
    if certainty_count:
        signals += 1
    if sentence_variation < 0.20:
        signals += 1
    if signals == 0:
        return "منخفض"
    if signals == 1:
        return "متوسط"
    return "مرتفع للمراجعة"


def clamp(value: float, lower: float = 0.0, upper: float = 1.0) -> float:
    return max(lower, min(upper, value))


def is_mostly_english(text: str) -> bool:
    letters = re.findall(r"[A-Za-z\u0600-\u06FF]", text)
    if not letters:
        return False
    english_letters = re.findall(r"[A-Za-z]", text)
    return len(english_letters) / len(letters) >= 0.70


def count_ai_vocabulary(text: str) -> Counter:
    lowered = text.lower()
    hits: Counter = Counter()
    for term in AI_VOCABULARY_TERMS:
        if re.search(r"^[A-Za-z -]+$", term):
            count = len(re.findall(rf"\b{re.escape(term.lower())}\b", lowered, flags=re.I))
        else:
            count = lowered.count(term.lower())
        if count:
            hits[term] = count
    return hits


def repeated_opener_ratio(sentences: Sequence[str]) -> float:
    if len(sentences) < 3:
        return 0.0
    openers = []
    for sentence in sentences:
        words = tokenize_words(sentence.lower())
        if words:
            openers.append(words[0])
    if not openers:
        return 0.0
    counts = Counter(openers)
    return max(counts.values()) / len(openers)


def current_section_label(line: str) -> Optional[str]:
    cleaned = re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", line.strip().lower()).strip(":：.- ")
    section_map = {
        "abstract": "abstract",
        "introduction": "introduction",
        "background": "introduction",
        "materials and methods": "methods",
        "methods": "methods",
        "methodology": "methods",
        "experimental setup": "methods",
        "results": "results",
        "discussion": "discussion",
        "conclusion": "conclusion",
        "conclusions": "conclusion",
        "references": "references",
        "bibliography": "references",
        "الملخص": "abstract",
        "المقدمة": "introduction",
        "المنهجية": "methods",
        "المواد والطرق": "methods",
        "النتائج": "results",
        "المناقشة": "discussion",
        "الخاتمة": "conclusion",
        "المراجع": "references",
    }
    return section_map.get(cleaned)


def exclusion_reason(line: str, active_section: Optional[str]) -> Optional[str]:
    stripped = line.strip()
    if not stripped:
        return None
    if active_section == "references":
        return "مراجع"
    if stripped.startswith(("|", "+---")) or stripped.count("|") >= 2:
        return "جداول"
    if re.search(r"^\s*(table|figure|fig\.|جدول|شكل)\s+\d+", stripped, flags=re.I):
        return "جداول وأشكال"
    if re.search(r"(```|^\s*(def|class|import|from|for|while|if|return)\b|^\s*(SELECT|CREATE|INSERT)\b)", stripped, flags=re.I):
        return "أكواد"
    if re.search(r"[{};]{2,}|(?:==|!=|<=|>=|->|=>)", stripped) and len(tokenize_words(stripped)) <= 18:
        return "أكواد"
    if re.search(r"[=∑√≈≤≥±×÷]", stripped) and len(tokenize_words(stripped)) <= 18:
        return "معادلات"
    tokens = tokenize_words(stripped)
    if tokens:
        numeric_tokens = sum(1 for token in tokens if re.search(r"\d", token))
        if numeric_tokens / len(tokens) >= 0.45:
            return "بيانات رقمية"
    if len(tokens) < 7:
        return "نص قصير غير مؤهل"
    return None


def isolate_qualifying_text(text: str) -> Tuple[str, Counter]:
    qualifying_lines: List[str] = []
    excluded: Counter = Counter()
    active_section: Optional[str] = None
    in_code_fence = False

    for line in text.splitlines():
        section = current_section_label(line)
        if section:
            active_section = section
            excluded["عناوين أقسام"] += len(line)
            continue
        if line.strip().startswith("```"):
            in_code_fence = not in_code_fence
            excluded["أكواد"] += len(line)
            continue
        if in_code_fence:
            excluded["أكواد"] += len(line)
            continue

        reason = exclusion_reason(line, active_section)
        if reason:
            excluded[reason] += len(line)
            continue
        qualifying_lines.append(line)

    qualifying_text = normalize_spacing("\n".join(qualifying_lines))
    long_sentences = [sentence for sentence in split_sentences(qualifying_text) if len(tokenize_words(sentence)) >= 7]
    return normalize_spacing(" ".join(long_sentences)), excluded


def academic_section_notes(text: str) -> Tuple[Tuple[str, ...], float]:
    sections: Counter = Counter()
    active_section: Optional[str] = None
    section_words: Counter = Counter()
    for line in text.splitlines():
        section = current_section_label(line)
        if section:
            sections[section] += 1
            active_section = section
            continue
        if active_section:
            section_words[active_section] += len(tokenize_words(line))

    notes: List[str] = []
    method_words = section_words.get("methods", 0)
    total_section_words = sum(section_words.values())
    method_ratio = method_words / total_section_words if total_section_words else 0.0
    if method_words:
        notes.append("تم رصد قسم منهجية/مواد؛ خُفف أثر انتظام الصياغة لأنه شائع في هذا القسم.")
    if sections.get("references"):
        notes.append("تم عزل قسم المراجع من حساب الأسلوب، مع إبقاء تنبيهات شكلية للمصداقية.")
    if not sections:
        notes.append("لم تظهر عناوين أقسام بحثية واضحة؛ جرى تحليل النص ككتلة نثرية واحدة.")
    return tuple(notes), clamp(method_ratio * 0.18, 0.0, 0.12)


def source_code_alerts(text: str) -> Tuple[str, ...]:
    alerts: List[str] = []
    fenced_blocks = len(re.findall(r"```[\s\S]*?```", text))
    if fenced_blocks:
        alerts.append(f"تم رصد {fenced_blocks} كتلة كود؛ عُزلت عن تحليل النثر وتحتاج مراجعة منفصلة.")
    if re.search(r"^\s*(def|class|import|from)\s+\w+", text, flags=re.M):
        alerts.append("توجد إشارات Python ظاهرة داخل المستند.")
    if re.search(r"\b(library|data\.frame|ggplot|<-)\b", text):
        alerts.append("توجد إشارات R أو تحليل بيانات داخل المستند.")
    if re.search(r"(?m)^\s{0,3}#{1,6}\s+\S|\*\*[^*\n]+\*\*|__[^_\n]+__|```", text):
        alerts.append("توجد آثار Markdown ظاهرة؛ راجعها إذا كان النص موجهاً لصيغة أكاديمية أو موسوعية لا تستخدم Markdown.")
    if re.search(r"\b(?:as an AI language model|as a large language model|I hope this helps|up to my last training update)\b", text, flags=re.I):
        alerts.append("توجد بقايا خطاب محادثة آلية داخل النص، مثل اعتذار النموذج أو حدود معرفته.")
    if re.search(r"\b(?:algorithm|pseudocode|خوارزمية)\b", text, flags=re.I):
        alerts.append("يوجد وصف خوارزمي؛ راجع توافقه مع المصادر أو المستودعات المفتوحة يدويًا.")
    return tuple(dict.fromkeys(alerts))


def translation_pattern_signal(text: str) -> float:
    words = tokenize_words(text)
    if len(words) < 40:
        return 0.0
    arabic = len(re.findall(r"[\u0600-\u06FF]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    total = arabic + latin
    mixed_script = min(arabic, latin) / total if total else 0.0
    literal_markers = count_pattern_matches(
        text,
        [
            r"\bin addition to that\b",
            r"\bfrom another side\b",
            r"\baccording to what preceded\b",
            r"\bthe matter\b",
            r"من جهة أخرى",
            r"بناء على ما سبق",
        ],
    )
    return clamp(mixed_script * 1.4 + min(0.35, literal_markers * 0.12))


def internal_overlap_signal(text: str) -> float:
    words = [word.lower() for word in tokenize_words(text)]
    if len(words) < 80:
        return 0.0
    ngrams = [" ".join(words[index : index + 6]) for index in range(len(words) - 5)]
    counts = Counter(ngrams)
    repeated = sum(count - 1 for count in counts.values() if count > 1)
    return clamp(repeated / max(12, len(ngrams)) * 5)


def citation_hygiene_alerts(text: str) -> Tuple[str, ...]:
    alerts: List[str] = []
    urls = re.findall(r"https?://\S+", text)
    dois = re.findall(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", text, flags=re.I)
    year_mentions = re.findall(r"\([^)]*\b(?:19|20)\d{2}\b[^)]*\)", text)
    bracket_citations = re.findall(r"\[[\d,\-; ]+\]", text)
    reference_lines = [
        line.strip()
        for line in text.splitlines()
        if re.search(r"\b(?:19|20)\d{2}\b", line) and len(tokenize_words(line)) >= 6
    ]
    if re.search(r"\b(?:et al\.?|وآخرون)\b", text, flags=re.I) and not year_mentions:
        alerts.append("توجد إحالة باسم مؤلفين دون سنة واضحة.")
    if re.search(r"\b(?:study|paper|research|دراسة|بحث)\b", text, flags=re.I) and not (year_mentions or bracket_citations or urls or dois):
        alerts.append("النص يذكر دراسات أو أبحاثًا دون مرجع ظاهر.")
    if urls and any(url.endswith((".", ",", ";")) for url in urls):
        alerts.append("بعض الروابط تنتهي بعلامات ترقيم؛ راجعها قبل التسليم.")
    if dois and any(len(doi) < 12 for doi in dois):
        alerts.append("يوجد DOI قصير أو غير مألوف ويحتاج تحققًا يدويًا.")
    if reference_lines and not (dois or urls):
        alerts.append("توجد مراجع بصياغة أكاديمية دون DOI أو رابط ظاهر؛ يلزم تحقق خارجي مثل Crossref أو PubMed.")
    if re.search(r"\b(?:Journal of Advanced|International Journal of Modern|Global Research in)\b", text, flags=re.I):
        alerts.append("ظهر اسم مجلة عام جدًا؛ راجعه يدويًا لأنه نمط شائع في المراجع المصطنعة.")
    return tuple(alerts)


def sentence_signal_score(sentence: str, avg_length: float, vocabulary_hits: Counter) -> SentenceSignal:
    words = tokenize_words(sentence)
    word_count = len(words)
    score = 0.0
    reasons: List[str] = []

    formulaic = count_pattern_matches(sentence, FORMULAIC_PATTERNS)
    if formulaic:
        score += 0.28
        reasons.append("عبارة قالبية")

    certainty = count_pattern_matches(sentence, CERTAINTY_PATTERNS)
    if certainty:
        score += 0.18
        reasons.append("جزم زائد")

    vocab_count = sum(count for term, count in count_ai_vocabulary(sentence).items())
    if vocab_count:
        score += min(0.24, 0.08 * vocab_count)
        reasons.append("مفردات شائعة في النصوص الآلية")

    if avg_length and 0.85 <= word_count / avg_length <= 1.15 and word_count >= 14:
        score += 0.12
        reasons.append("إيقاع شديد الانتظام")

    comma_count = sentence.count(",") + sentence.count("،")
    if word_count >= 24 and comma_count >= 2:
        score += 0.10
        reasons.append("جملة طويلة ومصقولة جدًا")

    if not reasons:
        reasons.append("إشارة منخفضة")

    score = clamp(score)
    if score >= 0.55:
        label = "مراجعة عالية"
    elif score >= 0.28:
        label = "مراجعة متوسطة"
    else:
        label = "مراجعة منخفضة"
    return SentenceSignal(sentence, score, label, tuple(reasons))


def compute_transparency_report(text: str) -> TransparencyReport:
    normalized = normalize_spacing(text)
    qualifying_text, excluded = isolate_qualifying_text(normalized)
    analysis_text = qualifying_text
    section_notes, section_discount = academic_section_notes(normalized)
    words = tokenize_words(analysis_text)
    sentences = split_sentences(analysis_text)
    sentence_lengths = [len(tokenize_words(sentence)) for sentence in sentences]
    avg_length = float(np.mean(sentence_lengths)) if sentence_lengths else 0.0
    variation = float(np.std(sentence_lengths) / avg_length) if avg_length else 0.0
    lexical_diversity = len(set(word.lower() for word in words)) / len(words) if words else 0.0
    formulaic_count = count_pattern_matches(analysis_text, FORMULAIC_PATTERNS)
    certainty_count = count_pattern_matches(analysis_text, CERTAINTY_PATTERNS)
    vocabulary_hits = count_ai_vocabulary(analysis_text)
    vocabulary_total = sum(vocabulary_hits.values())
    opener_ratio = repeated_opener_ratio(sentences)

    perplexity_signal = clamp((0.52 - lexical_diversity) * 1.7 + min(0.35, formulaic_count / 8))
    burstiness_signal = clamp((0.28 - variation) * 2.2) if len(sentences) >= 3 else 0.0
    paraphrase_signal = clamp((vocabulary_total / max(12, len(words))) * 4 + (0.20 if 0.22 <= variation <= 0.48 and vocabulary_total else 0))
    translation_signal = translation_pattern_signal(analysis_text)
    plagiarism_signal = internal_overlap_signal(analysis_text)

    esl_adjustment = 0.0
    if is_mostly_english(analysis_text) and avg_length <= 18 and formulaic_count == 0 and certainty_count == 0 and lexical_diversity >= 0.46:
        esl_adjustment = 0.14

    ai_generated_score = clamp(
        perplexity_signal * 0.40
        + burstiness_signal * 0.34
        + min(0.24, formulaic_count * 0.04)
        + min(0.16, certainty_count * 0.04)
        - esl_adjustment
        - section_discount
    )
    ai_paraphrased_score = clamp(
        paraphrase_signal * 0.58
        + translation_signal * 0.20
        + min(0.18, vocabulary_total / max(20, len(words)) * 2)
        - section_discount * 0.5
    )
    raw_score = (
        ai_generated_score * 0.42
        + ai_paraphrased_score * 0.30
        + min(0.20, certainty_count * 0.05)
        + clamp((opener_ratio - 0.34) * 0.8) * 0.16
        + plagiarism_signal * 0.08
        - esl_adjustment
        - section_discount
    )
    review_score = clamp(raw_score)
    false_positive_shield = review_score < 0.20

    if not normalized:
        classification = "لا يوجد نص"
    elif false_positive_shield:
        classification = "إشارة منخفضة / تحت عتبة 20%"
    elif review_score < 0.34:
        classification = "تحرير آلي محتمل"
    elif review_score < 0.52:
        classification = "مختلط محتمل"
    elif ai_paraphrased_score >= ai_generated_score:
        classification = "إعادة صياغة آلية محتملة"
    else:
        classification = "توليد آلي محتمل"

    sentence_signals = tuple(sentence_signal_score(sentence, avg_length, vocabulary_hits) for sentence in sentences)
    sorted_hits = tuple(vocabulary_hits.most_common(12))
    confidence = 0.0 if not analysis_text else review_score
    source_alerts = source_code_alerts(normalized)

    return TransparencyReport(
        chars=len(normalized),
        words=len(tokenize_words(normalized)),
        sentences=len(split_sentences(normalized)),
        qualifying_chars=len(analysis_text),
        excluded_chars=sum(excluded.values()),
        excluded_blocks=tuple(excluded.most_common()),
        minimum_ready=len(analysis_text) >= 250,
        classification=classification,
        confidence=confidence,
        ai_generated_score=ai_generated_score,
        ai_paraphrased_score=ai_paraphrased_score,
        plagiarism_signal=plagiarism_signal,
        perplexity_signal=perplexity_signal,
        burstiness_signal=burstiness_signal,
        paraphrase_signal=paraphrase_signal,
        translation_signal=translation_signal,
        esl_adjustment=esl_adjustment,
        false_positive_shield=false_positive_shield,
        ai_vocabulary_hits=sorted_hits,
        sentence_signals=sentence_signals,
        citation_alerts=citation_hygiene_alerts(normalized),
        source_code_alerts=source_alerts,
        section_notes=section_notes,
        model_update_note="مؤشر محلي قابل للتحديث: راجع قاموس الأنماط دوريًا، ويفضل كل 90 يومًا إذا تغيرت نماذج الكتابة أو سياسة المجلة.",
    )


def render_sentence_highlights(report: TransparencyReport) -> str:
    html_parts: List[str] = []
    for signal in report.sentence_signals:
        if signal.score >= 0.55:
            css_class = "sentence-high"
        elif signal.score >= 0.28:
            css_class = "sentence-medium"
        else:
            css_class = "sentence-low"
        title = html.escape("، ".join(signal.reasons))
        html_parts.append(
            f"<span class='{css_class}' title='{title}'>{html.escape(signal.text)}</span>"
        )
    return " ".join(html_parts)


def scan_distribution(report: TransparencyReport) -> Dict[str, int]:
    ai_share = clamp(max(report.ai_generated_score, report.ai_paraphrased_score) * 0.82 + report.confidence * 0.18)
    if report.false_positive_shield:
        ai_share = min(ai_share, 0.19)
    mixed_share = clamp(min(report.ai_generated_score, report.ai_paraphrased_score) * 0.55 + report.plagiarism_signal * 0.20)
    human_share = clamp(1.0 - ai_share - mixed_share)
    total = ai_share + mixed_share + human_share
    if total <= 0:
        return {"AI": 0, "Mixed": 0, "Human": 100}
    return {
        "AI": round(ai_share / total * 100),
        "Mixed": round(mixed_share / total * 100),
        "Human": max(0, 100 - round(ai_share / total * 100) - round(mixed_share / total * 100)),
    }


def scan_verdict_text(report: TransparencyReport) -> str:
    if report.false_positive_shield:
        return "الإشارات منخفضة وتحت عتبة الأمان؛ لا يوجد سبب كافٍ للتصعيد"
    if max(report.ai_generated_score, report.ai_paraphrased_score) >= 0.62:
        return "الإشارات المحلية مرتفعة وتستحق مراجعة بشرية دقيقة"
    if report.classification == "مختلط محتمل":
        return "النص يجمع بين إشارات طبيعية وإشارات آلية محتملة"
    if report.classification == "تحرير آلي محتمل":
        return "قد يكون النص بشريًا مع أثر تحرير أو انتظام زائد"
    return "لا تظهر إشارات محلية قوية، مع بقاء الحكم للمراجعة البشرية"


def render_advanced_scan_card(report: TransparencyReport, report_is_current: bool) -> str:
    distribution = scan_distribution(report)
    status = "المؤشرات محدّثة" if report_is_current else "اضغط المعالجة لتحديث المؤشرات"
    model_badge = "Offline heuristic review"
    dominant = max(distribution, key=distribution.get)
    dominant_class = {
        "AI": "scan-ai",
        "Mixed": "scan-mixed",
        "Human": "scan-human",
    }[dominant]
    return f"""
    <div class="advanced-scan">
        <div class="scan-heading">
            <div>
                <div class="scan-title">واقعية الفحص</div>
                <div class="scan-subtitle">DeepClean Authorship Signals <span>{model_badge}</span></div>
            </div>
            <div class="scan-status">{html.escape(status)}</div>
        </div>
        <div class="scan-body">
            <div class="scan-ring {dominant_class}">{dominant}</div>
            <div class="scan-verdict">
                <strong>{html.escape(scan_verdict_text(report))}</strong>
                <p>{report.qualifying_chars:,} qualifying characters · {report.words:,} words · {report.excluded_chars:,} excluded characters</p>
            </div>
        </div>
        <div class="scan-chips">
            <span class="chip chip-ai">إشارة آلية {distribution["AI"]}%</span>
            <span class="chip chip-mixed">إشارة مختلطة {distribution["Mixed"]}%</span>
            <span class="chip chip-human">إشارة بشرية {distribution["Human"]}%</span>
        </div>
        <div class="scan-toggle">
            <span>جمل تحتاج مراجعة</span>
            <span>مفردات مؤشرية</span>
        </div>
    </div>
    """


def external_detector_observations(report_text: str) -> List[str]:
    lowered = report_text.lower()
    observations: List[str] = []
    if "gptzero" in lowered:
        observations.append("المصدر المذكور يبدو GPTZero.")
    if "model 4.6b" in lowered:
        observations.append("النص يذكر إصدارًا/نموذجًا خارجيًا باسم Model 4.6b؛ يعرضه DeepClean كمعلومة منقولة فقط.")
    if "highly confident" in lowered:
        observations.append("التقرير الخارجي يستخدم صياغة ثقة عالية؛ لا يحولها DeepClean إلى حكم نهائي.")
    if "ai generated" in lowered or "ai-generated" in lowered:
        observations.append("التقرير الخارجي يدعي وجود تشابه مع نصوص مولدة آليًا.")
    if "compared" in lowered and "data" in lowered:
        observations.append("التقرير الخارجي يبرر النتيجة بالمقارنة مع بيانات مرجعية غير متاحة محليًا للتدقيق.")
    if not observations:
        observations.append("تم إرفاق تقرير خارجي غير مفسر آليًا؛ راجعه يدويًا مع المؤشرات المحلية.")
    return observations


def external_detector_severity(report_text: str) -> Tuple[str, str]:
    lowered = report_text.lower()
    high_markers = [
        "highly confident",
        "100%",
        "ai 100",
        "ai generated",
        "ai-generated",
        "was ai",
    ]
    medium_markers = ["mixed", "likely", "possibly", "may be", "similar to"]
    if any(marker in lowered for marker in high_markers):
        return (
            "مرتفع",
            "التقرير الخارجي يرفع أولوية المراجعة. لا ينبغي تجاهله حتى لو كانت المؤشرات المحلية أقل.",
        )
    if any(marker in lowered for marker in medium_markers):
        return (
            "متوسط",
            "التقرير الخارجي يحتوي على إشارة تحتاج مقارنة يدوية مع النص وسجل التحرير.",
        )
    return (
        "غير محدد",
        "التقرير الخارجي مرفق كسياق، لكن لا توجد فيه عبارة حاسمة يمكن تصنيفها محليًا.",
    )


def external_alignment_note(report_text: str, report: TransparencyReport) -> str:
    severity, _ = external_detector_severity(report_text)
    local_high = report.confidence >= 0.52 or max(report.ai_generated_score, report.ai_paraphrased_score) >= 0.52
    if severity == "مرتفع" and not local_high:
        return "يوجد تعارض مهم: الكاشف الخارجي مرتفع بينما مؤشرات DeepClean المحلية أقل. اعتمد هذا كإشارة فشل/نقص في التحليل المحلي، لا كنجاح."
    if severity == "مرتفع" and local_high:
        return "يوجد توافق عام: التقرير الخارجي والمؤشرات المحلية كلاهما يطلبان مراجعة بشرية دقيقة."
    if severity == "متوسط" and local_high:
        return "المؤشرات المحلية أعلى من التقرير الخارجي؛ راجع الجمل المظللة والمفردات قبل إعادة الفحص."
    return "لا يوجد تعارض واضح، لكن التقرير الخارجي يبقى قرينة منفصلة لا تتحول إلى حكم نهائي."


def render_external_detector_card(report_text: str, report: Optional[TransparencyReport] = None) -> str:
    cleaned = report_text.strip()
    if not cleaned:
        return ""
    severity, severity_note = external_detector_severity(cleaned)
    alignment = external_alignment_note(cleaned, report) if report is not None else ""
    observations = "".join(f"<li>{html.escape(item)}</li>" for item in external_detector_observations(cleaned))
    alignment_html = f"<p><strong>المقارنة مع DeepClean:</strong> {html.escape(alignment)}</p>" if alignment else ""
    return f"""
    <div class="external-report">
        <div class="external-title">تقرير كاشف خارجي مرفق</div>
        <p>هذا النص منقول من أداة خارجية ولا ينتجه DeepClean. يُستخدم كقرينة مراجعة فقط.</p>
        <p><strong>مستوى التحذير الخارجي:</strong> {html.escape(severity)} — {html.escape(severity_note)}</p>
        {alignment_html}
        <blockquote>{html.escape(cleaned)}</blockquote>
        <ul>{observations}</ul>
    </div>
    """


def compute_stats(original: str, revised: str) -> RevisionStats:
    original_words = [word.lower() for word in tokenize_words(original)]
    revised_words = [word.lower() for word in tokenize_words(revised)]
    revised_sentences = split_sentences(revised)
    sentence_lengths = [len(tokenize_words(sentence)) for sentence in revised_sentences]
    lexical_diversity = len(set(revised_words)) / len(revised_words) if revised_words else 0.0
    similarity = jaccard_similarity(original_words, revised_words)
    formulaic_count = count_pattern_matches(revised, FORMULAIC_PATTERNS)
    certainty_count = count_pattern_matches(revised, CERTAINTY_PATTERNS)
    avg_length = float(np.mean(sentence_lengths)) if sentence_lengths else 0.0
    sentence_variation = float(np.std(sentence_lengths) / avg_length) if avg_length else 0.0
    return RevisionStats(
        words_original=len(original_words),
        words_revised=len(revised_words),
        sentences_revised=len(revised_sentences),
        avg_sentence_length=avg_length,
        lexical_diversity=lexical_diversity,
        similarity=similarity,
        formulaic_phrase_count=formulaic_count,
        certainty_marker_count=certainty_count,
        sentence_length_variation=sentence_variation,
        authorship_review_band=authorship_review_band(formulaic_count, certainty_count, sentence_variation),
    )


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


def process_text_callback() -> None:
    text = (
        st.session_state.get("text_input", "").strip()
        or st.session_state.get("paste_area", "").strip()
    )
    if not text:
        st.warning("أدخل نصًا أو ارفع ملفًا قبل المعالجة.")
        return
    st.session_state.text_input = text

    if is_mostly_arabic(text):
        engine = ArabicEditorialRevisionEngine(
            intensity=st.session_state.get("intensity", 2),
            text=text,
            preserve_word_count=st.session_state.get("preserve_word_count", True),
        )
    else:
        engine = AcademicRevisionEngine(
            domain=st.session_state.get("domain", "general"),
            intensity=st.session_state.get("intensity", 2),
            text=text,
            preserve_word_count=st.session_state.get("preserve_word_count", True),
        )
    revised = polish_layout_artifacts(engine.run())
    st.session_state.revised_text = revised
    st.session_state.stats = compute_stats(text, revised)
    st.session_state.transparency_report = compute_transparency_report(text)
    st.session_state.processed_source_text = text
    st.session_state.processing_done = True


defaults = {
    "revised_text": "",
    "processing_done": False,
    "text_input": "",
    "domain": "general",
    "intensity": 2,
    "preserve_word_count": True,
    "stats": RevisionStats(0, 0, 0, 0.0, 0.0, 1.0, 0, 0, 0.0, "منخفض"),
    "transparency_report": compute_transparency_report(""),
    "processed_source_text": "",
    "external_detector_report": "",
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


st.title("DeepClean Studio")
st.caption("محرر أكاديمي لمراجعة الوضوح والتدفق مع الحفاظ على المعنى والمراجع.")
st.markdown(
    """
    <style>
    .transparency-panel {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 16px;
        background: #fbfbfc;
        line-height: 1.85;
    }
    .sentence-high,
    .sentence-medium,
    .sentence-low {
        border-radius: 6px;
        padding: 2px 4px;
        margin: 2px 1px;
        display: inline;
    }
    .sentence-high { background: #fed7aa; border-bottom: 2px solid #f97316; }
    .sentence-medium { background: #fef3c7; border-bottom: 2px solid #f59e0b; }
    .sentence-low { background: #dcfce7; border-bottom: 2px solid #22c55e; }
    .signal-note {
        color: #475569;
        font-size: 0.92rem;
    }
    .advanced-scan {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 18px;
        background: #ffffff;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
        margin-bottom: 16px;
    }
    .scan-heading,
    .scan-body,
    .scan-chips,
    .scan-toggle {
        display: flex;
        align-items: center;
        gap: 14px;
        flex-wrap: wrap;
    }
    .scan-heading {
        justify-content: space-between;
        padding-bottom: 14px;
        border-bottom: 1px solid #e5e7eb;
    }
    .scan-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #111827;
    }
    .scan-subtitle {
        color: #475569;
        margin-top: 4px;
    }
    .scan-subtitle span,
    .scan-status {
        background: #e5e7eb;
        border-radius: 6px;
        padding: 3px 8px;
        color: #334155;
        font-size: 0.85rem;
    }
    .scan-body {
        padding: 18px 0 12px;
    }
    .scan-ring {
        width: 86px;
        height: 86px;
        border-radius: 50%;
        display: grid;
        place-items: center;
        font-weight: 800;
        background: #fff;
    }
    .scan-ai { border: 6px solid #eab308; color: #92400e; }
    .scan-mixed { border: 6px solid #22c55e; color: #166534; }
    .scan-human { border: 6px solid #10b981; color: #047857; }
    .scan-verdict strong {
        display: block;
        font-size: 1.16rem;
        color: #1f2937;
        margin-bottom: 6px;
    }
    .scan-verdict p {
        margin: 0;
        color: #64748b;
    }
    .chip {
        border: 2px solid;
        border-radius: 999px;
        padding: 7px 13px;
        font-weight: 700;
        background: #fff;
    }
    .chip-ai { border-color: #eab308; color: #3f3f46; }
    .chip-mixed { border-color: #86efac; color: #64748b; }
    .chip-human { border-color: #6ee7b7; color: #64748b; }
    .scan-toggle {
        margin-top: 18px;
        background: #f1f5f9;
        border-radius: 999px;
        padding: 6px;
        justify-content: center;
    }
    .scan-toggle span {
        min-width: 150px;
        text-align: center;
        background: #fff;
        border-radius: 999px;
        padding: 9px 16px;
        color: #111827;
        font-weight: 700;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.08);
    }
    .external-report {
        border: 1px solid #dbeafe;
        border-radius: 8px;
        padding: 14px 16px;
        background: #eff6ff;
        margin: 14px 0;
        color: #1e3a8a;
    }
    .external-title {
        font-weight: 800;
        color: #1e40af;
        margin-bottom: 6px;
    }
    .external-report blockquote {
        margin: 10px 0;
        padding: 10px 12px;
        border-left: 4px solid #60a5fa;
        background: #ffffff;
        color: #334155;
    }
    .external-report ul {
        margin-bottom: 0;
    }
    .academic-preview {
        border: 1px solid #d7dee8;
        border-radius: 8px;
        background: #ffffff;
        color: #111827;
        padding: 18px 20px;
        margin-bottom: 12px;
        font-family: "Times New Roman", Arial, serif;
        line-height: 1.55;
    }
    .academic-preview h3 {
        font-size: 1.05rem;
        margin: 14px 0 8px;
        color: #111827;
        font-weight: 700;
    }
    .academic-preview h3:first-child {
        margin-top: 0;
    }
    .academic-preview p {
        margin: 0 0 9px;
        text-align: justify;
    }
    .academic-preview p[dir="rtl"] {
        text-align: right;
    }
    .academic-preview ul {
        margin: 0 0 10px 22px;
        padding: 0;
    }
    .academic-preview table {
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0;
        font-size: 0.95rem;
    }
    .academic-preview td {
        border: 1px solid #cbd5e1;
        padding: 6px 8px;
        vertical-align: top;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("الإعدادات")
    input_option = st.radio("مصدر النص", ("رفع ملف", "لصق نص"), key="source_radio")

    text_input = ""
    if input_option == "رفع ملف":
        uploaded_file = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"], key="file_uploader")
        if uploaded_file is not None:
            try:
                text_input = extract_uploaded_text(uploaded_file)
            except Exception as exc:
                st.error(f"تعذر قراءة الملف: {exc}")
    else:
        text_input = st.text_area("ألصق النص الأكاديمي هنا", height=220, key="paste_area")

    if text_input:
        st.session_state.text_input = text_input

    st.slider(
        "قوة المراجعة",
        1,
        5,
        2,
        help="1=تنظيف خفيف، 5=مراجعة أسلوبية أوسع مع الحفاظ على المعنى.",
        key="intensity",
    )
    st.selectbox(
        "المجال الأكاديمي",
        ("medical", "engineering", "humanities", "general"),
        format_func=lambda value: {
            "medical": "طبي",
            "engineering": "هندسي",
            "humanities": "علوم إنسانية",
            "general": "عام",
        }[value],
        key="domain",
    )
    st.checkbox("حافظ على حجم النص قدر الإمكان", key="preserve_word_count")
    st.text_area(
        "تقرير كاشف خارجي (اختياري)",
        key="external_detector_report",
        height=120,
        help="ألصق هنا نتيجة GPTZero أو غيره إن وجدت. سيعرضها التطبيق كمرفق خارجي لا كحكم صادر منه.",
    )

    st.button(
        "بدء المراجعة",
        type="primary",
        use_container_width=True,
        on_click=process_text_callback,
    )

    st.markdown("---")
    st.subheader("مؤشرات الجودة")
    stats: RevisionStats = st.session_state.stats
    st.metric("الكلمات قبل/بعد", f"{stats.words_original} / {stats.words_revised}")
    st.metric("متوسط طول الجملة", f"{stats.avg_sentence_length:.1f}")
    st.metric("التنوع اللفظي", f"{stats.lexical_diversity:.2f}")
    st.metric("تشابه المفردات", f"{stats.similarity:.2f}")
    st.info("هذه المقاييس تساعد على المراجعة فقط، ولا تمثل حكمًا على الأصالة أو القبول الأكاديمي.")

    st.metric("عبارات قالبية", stats.formulaic_phrase_count)
    st.metric("عبارات جزم زائد", stats.certainty_marker_count)
    st.metric("تباين أطوال الجمل", f"{stats.sentence_length_variation:.2f}")
    st.metric("نطاق المراجعة", stats.authorship_review_band)
    st.info(
        "درجات الكشف الآلي تعبر عن ثقة النموذج في تصنيف الوثيقة بأكملها، وليست نسبة الجمل المكتوبة آليًا. "
        "هذه المؤشرات محلية للمراجعة التحريرية والحوار، ولا تمثل حكمًا نهائيًا على الأصالة أو القبول الأكاديمي."
    )

    if st.session_state.revised_text:
        word_file = create_word_document(st.session_state.revised_text)
        st.download_button(
            "تنزيل النص المراجع المنسق (Word)",
            data=word_file,
            file_name="deepclean_revised.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )


col1, col2 = st.columns(2)
with col1:
    st.subheader("النص الأصلي")
    if st.session_state.text_input:
        st.text_area(
            "النص الأصلي",
            st.session_state.text_input,
            height=420,
            label_visibility="collapsed",
            key="original_text_display",
        )
        st.caption(f"عدد الكلمات: {len(tokenize_words(st.session_state.text_input))}")
    else:
        st.info("ارفع ملفًا أو ألصق نصًا من الشريط الجانبي.")

with col2:
    st.subheader("النص المراجع")
    if st.session_state.revised_text:
        st.markdown(render_academic_preview(st.session_state.revised_text), unsafe_allow_html=True)
        st.text_area(
            "النص المراجع",
            st.session_state.revised_text,
            height=420,
            label_visibility="collapsed",
            key="revised_text_display",
        )
        st.caption(f"عدد الكلمات: {len(tokenize_words(st.session_state.revised_text))}")
    else:
        st.info("ستظهر النسخة المراجعة هنا بعد المعالجة.")


if st.session_state.text_input:
    report_is_current = (
        st.session_state.processing_done
        and st.session_state.processed_source_text == st.session_state.text_input.strip()
    )
    report: TransparencyReport = (
        st.session_state.transparency_report
        if report_is_current
        else compute_transparency_report(st.session_state.text_input)
    )
    st.markdown("---")
    with st.expander("لوحة الشفافية والتحليل السبعي", expanded=report_is_current):
        if not report.minimum_ready:
            st.warning(
                f"النص المؤهل للفحص يحتوي على {report.qualifying_chars} حرفًا بعد عزل المراجع والجداول والأكواد. يفضل توفر 250 حرفًا مؤهلًا على الأقل قبل الاعتماد على إشارات المراجعة."
            )

        overview_tab, sentence_tab, vocabulary_tab, support_tab = st.tabs(
            ["التصنيف", "تظليل الجمل", "المفردات", "الدعم والتحقق"]
        )

        with overview_tab:
            st.markdown(render_advanced_scan_card(report, report_is_current), unsafe_allow_html=True)
            if not report_is_current:
                st.info("غيّرت النص بعد آخر معالجة. اضغط بدء المراجعة لتحديث المؤشرات والنص المراجع معًا.")
            metric_cols = st.columns(5)
            metric_cols[0].metric("نطاق المؤشرات", report.classification)
            metric_cols[1].metric("قوة الإشارة", f"{report.confidence * 100:.0f}%")
            metric_cols[2].metric("الحروف", report.chars)
            metric_cols[3].metric("الكلمات", report.words)
            metric_cols[4].metric("الجمل", report.sentences)
            st.progress(report.confidence)
            if report.false_positive_shield:
                st.success("درع عتبة 20% فعّال: الإشارات منخفضة، لذلك لا تظهر النتيجة كاتهام أو حكم أصالة.")
            st.caption(
                "هذه قراءة محلية قابلة للتفسير. لا تثبت التأليف الآلي، ولا تنفيه، ولا تغني عن مراجعة بشرية وسجل تحرير ومصادر قابلة للتحقق."
            )
            external_report = st.session_state.get("external_detector_report", "").strip()
            if external_report:
                st.markdown(render_external_detector_card(external_report, report), unsafe_allow_html=True)

            q_cols = st.columns(3)
            q_cols[0].metric("النص المؤهل للفحص", f"{report.qualifying_chars} حرف")
            q_cols[1].metric("المستبعد من الفحص", f"{report.excluded_chars} حرف")
            q_cols[2].metric("تحديث الأنماط", "توصية يدوية")

            breakdown_rows = pd.DataFrame(
                [
                    {"الفئة": "تشابه مع نص مولد آليًا", "الدرجة": f"{report.ai_generated_score:.2f}"},
                    {"الفئة": "اشتباه إعادة صياغة آلية", "الدرجة": f"{report.ai_paraphrased_score:.2f}"},
                    {"الفئة": "تكرار داخلي", "الدرجة": f"{report.plagiarism_signal:.2f}"},
                    {"الفئة": "أثر ترجمة أو انتقال لغوي", "الدرجة": f"{report.translation_signal:.2f}"},
                ]
            )
            st.dataframe(breakdown_rows, use_container_width=True, hide_index=True)

            signal_rows = pd.DataFrame(
                [
                    {"المكون": "الحيرة/تنوع المفردات", "القيمة": f"{report.perplexity_signal:.2f}"},
                    {"المكون": "التدفق/تباين الجمل", "القيمة": f"{report.burstiness_signal:.2f}"},
                    {"المكون": "درع إعادة الصياغة", "القيمة": f"{report.paraphrase_signal:.2f}"},
                    {"المكون": "تعديل ESL", "القيمة": f"-{report.esl_adjustment:.2f}"},
                ]
            )
            st.dataframe(signal_rows, use_container_width=True, hide_index=True)
            if report.excluded_blocks:
                excluded_rows = pd.DataFrame(
                    [{"العنصر المعزول": reason, "حروف مستبعدة": count} for reason, count in report.excluded_blocks]
                )
                st.dataframe(excluded_rows, use_container_width=True, hide_index=True)

        with sentence_tab:
            if report.sentence_signals:
                highlighted = render_sentence_highlights(report)
                st.markdown(f"<div class='transparency-panel'>{highlighted}</div>", unsafe_allow_html=True)
                st.caption("البرتقالي = مراجعة عالية، الأصفر = مراجعة متوسطة، الأخضر = إشارة منخفضة.")
                sentence_rows = pd.DataFrame(
                    [
                        {
                            "الجملة": signal.text,
                            "النطاق": signal.label,
                            "الدرجة": f"{signal.score:.2f}",
                            "السبب": "، ".join(signal.reasons),
                        }
                        for signal in report.sentence_signals
                    ]
                )
                st.dataframe(sentence_rows, use_container_width=True, hide_index=True)
            else:
                st.info("لا توجد جمل كافية للتظليل بعد.")

        with vocabulary_tab:
            if report.ai_vocabulary_hits:
                vocab_rows = pd.DataFrame(
                    [{"المفردة": term, "عدد الظهور": count} for term, count in report.ai_vocabulary_hits]
                )
                st.dataframe(vocab_rows, use_container_width=True, hide_index=True)
            else:
                st.success("لم تظهر مفردات آلية متكررة ضمن القاموس المحلي.")
            st.caption("القاموس المحلي يركز على إشارات تحريرية شائعة، وقد لا يلتقط كل الأنماط أو اللغات.")

        with support_tab:
            component_rows = pd.DataFrame(
                [{"المكون": title, "كيف يستخدمه DeepClean": description} for title, description in TRANSPARENCY_COMPONENTS]
            )
            st.dataframe(component_rows, use_container_width=True, hide_index=True)
            st.subheader("حدود الدقة")
            limitations_rows = pd.DataFrame(
                [{"البند": title, "المعنى": description} for title, description in REALISM_LIMITATIONS]
            )
            st.dataframe(limitations_rows, use_container_width=True, hide_index=True)
            external_report = st.session_state.get("external_detector_report", "").strip()
            if external_report:
                st.markdown(render_external_detector_card(external_report, report), unsafe_allow_html=True)
            for note in report.section_notes:
                st.info(note)
            if report.source_code_alerts:
                for alert in report.source_code_alerts:
                    st.warning(alert)
            if report.citation_alerts:
                for alert in report.citation_alerts:
                    st.warning(alert)
            else:
                st.success("لم تظهر تنبيهات مرجعية محلية واضحة.")
            st.info(
                "فحص الهلوسة هنا محلي ومبدئي: يراجع شكل الإحالات والروابط فقط. التحقق الكامل من المصادر يحتاج اتصالًا بقاعدة بيانات أو بحثًا موثقًا."
            )
            st.info(
                "تقرير Writing Replay يحتاج تكاملًا مع سجل تحرير خارجي مثل Google Docs؛ التطبيق يجهز مكانه المفاهيمي دون تتبع كتابة المستخدم."
            )
            st.caption(report.model_update_note)


if st.session_state.revised_text and st.session_state.text_input:
    st.markdown("---")
    with st.expander("تفاصيل التغييرات", expanded=False):
        diff_html = word_level_diff(st.session_state.text_input, st.session_state.revised_text)
        st.markdown(
            """
            <style>
            .diff-box {background:#fafafa;border:1px solid #e5e7eb;padding:14px;border-radius:8px;line-height:1.8}
            .removed {color:#b91c1c;text-decoration:line-through}
            .added {color:#047857;font-weight:600}
            </style>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(f"<div class='diff-box'>{diff_html}</div>", unsafe_allow_html=True)

        original_set = Counter(word.lower() for word in tokenize_words(st.session_state.text_input))
        revised_set = Counter(word.lower() for word in tokenize_words(st.session_state.revised_text))
        removed = sorted((original_set - revised_set).elements())
        added = sorted((revised_set - original_set).elements())
        if removed or added:
            changes = pd.DataFrame(
                {
                    "النوع": ["محذوف"] * len(removed) + ["مضاف"] * len(added),
                    "الكلمة": removed + added,
                }
            )
            st.dataframe(changes, use_container_width=True, hide_index=True)
        else:
            st.info("لا توجد تغييرات جوهرية على مستوى الكلمات.")

st.markdown("---")
st.caption("DeepClean Studio © 2026 - للمراجعة التعليمية والبحثية المسؤولة.")
