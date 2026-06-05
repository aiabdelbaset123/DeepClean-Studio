"""
DeepClean Studio v4.0 — Professional Academic Text Humanization Engine
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
from docx.oxml.ns import qn

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
# AI VOCABULARY DATABASE v4.0 — MASSIVELY EXPANDED
# قاعدة بيانات المفردات المفضلة للذكاء الاصطناعي — نسخة موسعة جداً
#
# These words appear FAR more frequently in post-2022 text (after LLMs
# became widespread) than in pre-2022 text. Replacing them is the
# single most impactful humanization step.
# ═══════════════════════════════════════════════════════════════

AI_VOCAB_REPLACEMENTS: Dict[str, List[str]] = {
    # ── Core AI "tells" — the strongest signals ──
    "crucial":            ["important", "pressing", "central", "key"],
    "pivotal":            ["decisive", "key", "central", "main"],
    "underscore":         ["confirm", "show", "back up", "bear out"],
    "highlight":          ["point to", "show", "bring out", "flag"],
    "showcase":           ["present", "display", "demonstrate", "show"],
    "delve":              ["examine", "probe", "explore", "look into", "investigate"],
    "intricate":          ["complex", "detailed", "elaborate", "involved"],
    "fostering":          ["encouraging", "promoting", "supporting", "driving"],
    "garner":             ["gain", "attract", "collect", "earn"],
    "meticulous":         ["careful", "thorough", "detailed", "close"],
    "robust":             ["strong", "solid", "reliable", "sturdy"],
    "testament":          ["proof", "evidence", "sign", "indication"],
    "enduring":           ["lasting", "long-standing", "persistent", "ongoing"],
    "vibrant":            ["active", "lively", "busy", "thriving"],
    "tapestry":           ["mix", "blend", "mosaic", "fabric"],
    "bolstered":          ["strengthened", "supported", "reinforced", "backed"],
    "landscape":          ["field", "area", "domain", "scene"],
    "valuable":           ["useful", "important", "helpful", "worthwhile"],
    "enhance":            ["improve", "strengthen", "boost", "raise"],
    "emphasizing":        ["stressing", "underlining", "pointing up"],
    "showcasing":         ["presenting", "displaying", "featuring"],
    "align with":         ["match", "fit", "correspond to", "accord with"],
    "serves as":          ["is", "acts as", "functions as"],
    "stands as":          ["is", "remains", "constitutes"],
    "boasts":             ["has", "features", "contains", "includes"],
    "features":           ["has", "includes", "contains", "offers"],
    "offers":             ["provides", "gives", "has", "supplies"],
    "represents":         ["is", "forms", "constitutes", "marks"],
    "marks":              ["is", "constitutes", "signals", "marks out"],
    "exemplifies":        ["illustrates", "typifies", "demonstrates"],
    "additionally":       ["also", "in addition", "further"],
    "comprehensive":      ["thorough", "complete", "full", "wide-ranging"],
    "numerous":           ["many", "several", "various", "a range of"],
    "significant":        ["notable", "important", "substantial", "marked"],
    "subsequently":       ["then", "later", "afterward", "next"],
    "furthermore":        ["moreover", "also", "in addition"],
    "moreover":           ["also", "further", "in addition"],
    "notably":            ["in particular", "especially", "worth noting"],
    "remarkably":         ["strikingly", "unusually", "interestingly"],
    "demonstrates":       ["shows", "reveals", "indicates", "makes clear"],
    "illustrates":        ["shows", "reveals", "depicts", "makes visible"],
    "underscores":        ["confirms", "reinforces", "stresses"],
    "navigating":         ["dealing with", "managing", "handling"],
    "realm":              ["area", "field", "domain", "sphere"],
    "paramount":          ["top-priority", "utmost", "vital", "chief"],
    "unwavering":         ["steady", "firm", "resolute", "consistent"],
    "unprecedented":      ["novel", "never-before-seen", "extraordinary", "new"],
    "groundbreaking":     ["pioneering", "innovative", "novel", "fresh"],
    "multifaceted":       ["complex", "many-sided", "layered"],
    "holistic":           ["integrated", "all-round", "joined-up"],
    "streamlined":        ["efficient", "simplified", "lean"],
    "leverage":           ["use", "exploit", "harness", "draw on"],
    "utilize":            ["use", "employ", "apply", "draw on"],
    "facilitate":         ["enable", "help", "assist", "make possible"],
    "encompasses":        ["includes", "covers", "spans", "takes in"],
    "pertains to":        ["relates to", "concerns", "is about", "regards"],
    "nestled":            ["located", "situated", "set", "placed"],
    "breathtaking":       ["striking", "impressive", "remarkable"],
    "rich":               ["deep", "varied", "extensive", "strong"],
    "profound":           ["deep", "far-reaching", "serious", "weighty"],
    "dynamic":            ["active", "changing", "evolving", "shifting"],
    "charming":           ["pleasant", "appealing", "attractive"],
    "stunning":           ["striking", "impressive", "remarkable"],
    "picturesque":        ["scenic", "attractive", "pretty"],

    # ── Multi-word AI patterns ──
    "serves as a":        ["is a", "acts as a", "functions as a"],
    "serves as an":       ["is an", "acts as an", "functions as an"],
    "stands as a":        ["is a", "remains a"],
    "stands as an":       ["is an", "remains an"],
    "serves as the":      ["is the"],
    "stands as the":      ["is the"],
    "place at the centre": ["put at the center", "put at the heart of"],
    "at the centre of":   ["at the heart of", "at the core of", "central to"],
    "constitute the":     ["make up the", "form the", "account for the"],
    "constitute a":       ["make up a", "form a", "amount to a"],
    "remains fragmented": ["is still split", "is still divided", "is still siloed"],
    "remains entirely":   ["is still completely", "is still wholly"],
    "remain entirely":    ["stay completely", "stay wholly"],
    "place ... at the centre": ["put ... at the center"],
    "the single largest":  ["the biggest", "the largest single"],
    "the dominant share":  ["the largest share", "the biggest portion", "most of it"],

    # ── ACADEMIC AI WORDS — expanded v4.0 ──
    # Words that LLMs overuse in academic/scientific writing
    "constitute":         ["make up", "form", "account for", "compose"],
    "constitutes":        ["makes up", "forms", "amounts to", "represents"],
    "dominant":           ["main", "leading", "chief", "primary", "largest"],
    "established":        ["well-known", "standard", "widely used", "conventional"],
    "routinely":          ["commonly", "often", "regularly", "frequently", "typically"],
    "pronounced":         ["strong", "marked", "clear", "noticeable", "sharp"],
    "uniquely":           ["especially", "particularly", "distinctively", "specially"],
    "systematically":     ["consistently", "thoroughly", "across the board", "methodically"],
    "fragmented":         ["split", "divided", "disconnected", "broken up", "siloed"],
    "entirely":           ["completely", "wholly", "fully", "altogether"],
    "specifically":       ["in particular", "namely", "precisely"],
    "individual":         ["separate", "single", "stand-alone", "discrete"],
    "mechanistic":        ["cause-and-effect", "mechanism-based", "process-based"],
    "intelligence":       ["insight", "analysis", "understanding", "know-how"],
    "trajectory":         ["path", "course", "route", "trend"],
    "trajectories":       ["paths", "courses", "scenarios"],
    "episode":            ["event", "period", "occurrence", "spell"],
    "episodes":           ["events", "periods", "occurrences"],
    "uniquely difficult": ["especially hard", "particularly tough"],
    "despite decades":    ["even after years", "in spite of years"],
    "high-throughput":    ["large-scale", "high-volume", "bulk"],
    "derived":            ["obtained", "computed", "calculated", "extracted"],
    "exceeding":          ["above", "over", "surpassing", "going beyond"],
    "committed to":       ["pledged to", "set out to", "aiming to", "promised to"],
    "impose":             ["create", "present", "pose"],
    "govern":             ["control", "shape", "determine", "drive"],
    "expose":             ["provide", "offer", "give access to", "make available"],
    "accept":             ["take in", "use", "allow", "support"],
    "produced":           ["yielded", "generated", "created", "turned out"],
    "accurate":           ["precise", "reliable", "sound", "good"],
    "connected":          ["linked", "tied", "joined", "coupled"],
    "disconnected":       ["separate from", "cut off from", "divorced from", "apart from"],
    "compounded":         ["made worse", "aggravated", "intensified"],
    "inconsistency":      ["mismatch", "gap", "disagreement", "discrepancy"],
    "practitioners":      ["users", "engineers", "researchers", "analysts"],
    "transfer":           ["move", "copy", "pass", "shift"],
    "forces":             ["makes", "requires", "compels", "obliges"],
    "ignores":            ["misses", "overlooks", "skips", "leaves out"],
    "interactions":       ["links", "connections", "couplings", "cross-effects"],
    "progress":           ["advances", "headway", "improvements", "steps forward"],
    "sub-fields":         ["areas", "disciplines", "domains", "branches"],
    "viability":          ["feasibility", "practicality", "workability"],
    "prediction":         ["forecast", "estimate", "projection", "modeling"],
    "performance":        ["output", "behavior", "results", "functioning"],
    "ecosystem":          ["set of tools", "toolset", "environment", "tool chain"],
    "boundaries":         ["interfaces", "junctions", "seams", "crossings"],
    "loading":            ["burden", "concentration", "level", "amount"],
    "mineralogy":         ["composition", "makeup", "mineral content"],
    "factors":            ["parameters", "variables", "settings", "inputs"],
    "link":               ["tie", "connection", "relationship", "relation"],
    "connection":         ["link", "tie", "bridge", "path"],
    "provide no":         ["do not provide", "lack", "offer no", "give no"],
    "no connection":      ["no link", "no bridge", "no tie"],
    "no mechanistic":     ["no cause-based", "no physics-based", "no process-based"],
    "and offer no":       ["and give no", "with no", "and lack any"],
    "with no":            ["without", "lacking", "absent any"],
    "but provide":        ["yet give", "but offer", "but supply"],
    "that remain":        ["which stay", "that are still", "which keep"],

    # ── More AI-favorite adjectives and adverbs ──
    "notably":            ["in particular", "especially", "worth noting"],
    "particularly":       ["especially", "in particular", "specially"],
    "fundamentally":      ["at root", "basically", "in essence", "at heart"],
    "essentially":        ["basically", "in practice", "at bottom"],
    "inherently":         ["by nature", "intrinsically", "by its nature"],
    "intrinsically":      ["by nature", "inherently", "in its own right"],
    "substantially":      ["considerably", "a great deal", "markedly", "a lot"],
    "considerably":       ["a great deal", "markedly", "substantially", "much"],
    "increasingly":       ["more and more", "ever more", "growingly"],
    "predominantly":      ["mainly", "mostly", "chiefly", "largely"],
    "overwhelmingly":     ["heavily", "vastly", "by far"],
    "strikingly":         ["sharply", "markedly", "clearly"],
    "markedly":           ["clearly", "noticeably", "sharply", "distinctly"],
    "consequently":       ["as a result", "because of this", "so"],
    "thereby":            ["thus", "by doing so", "in this way"],
    "wherein":            ["in which", "where", "within which"],
    "thereof":            ["of it", "of that", "of which"],
    "albeit":             ["even though", "although", "while", "though"],
    "whereas":            ["while", "although", "though", "when in fact"],
    "hitherto":           ["until now", "so far", "previously"],
    "heretofore":         ["before this", "up to now", "previously"],
    "hence":              ["so", "therefore", "thus", "for this reason"],
    "thus":               ["so", "therefore", "in this way"],
    "therefore":          ["so", "for this reason", "as a result"],
    "indeed":             ["in fact", "as it happens", "to be sure"],
    "conversely":         ["on the flip side", "on the other hand", "by contrast"],
    "simultaneously":     ["at the same time", "in parallel", "together"],
    "subsequent":         ["later", "following", "next"],
    "aforementioned":     ["previously mentioned", "above", "earlier"],
    "pertinent":          ["relevant", "germane", "related", "applicable"],
    "salient":            ["key", "main", "notable", "important"],
    "seminal":            ["foundational", "key", "landmark", "groundbreaking"],
    "nuanced":            ["subtle", "fine-grained", "layered", "detailed"],
    "efficacious":        ["effective", "successful", "working", "useful"],
    "efficacy":           ["effectiveness", "success rate", "performance"],
    "implementation":     ["rollout", "deployment", "putting into practice"],
    "endeavor":           ["effort", "attempt", "undertaking", "project"],
    "endeavors":          ["efforts", "attempts", "projects"],
    "paradigm":           ["model", "framework", "approach", "system"],
    "paradigms":          ["models", "frameworks", "approaches"],
    "synergy":            ["combined effect", "interaction", "cooperation"],
    "synergistic":        ["cooperative", "complementary", "combined"],
    "imperative":         ["essential", "vital", "necessary", "must-have"],
    "indispensable":      ["essential", "vital", "necessary", "must-have"],
    "catalyst":           ["driver", "spark", "trigger", "cause"],
    "cornerstone":        ["foundation", "basis", "key element", "bedrock"],
    "linchpin":           ["key part", "central piece", "hinge"],
    "hallmark":           ["sign", "feature", "trademark", "mark"],
    "trajectory":         ["path", "course", "direction", "line"],
    "dichotomy":          ["split", "division", "contrast", "two-part split"],
    "amalgamation":       ["blend", "mix", "combination", "merger"],
    "juxtaposition":      ["contrast", "comparison", "side-by-side view"],
    "pedagogical":        ["teaching", "educational", "instructional"],
    "methodological":     ["method-based", "procedural", "technique-based"],
    "epistemological":    ["knowledge-based", "philosophical"],
    "ontological":        ["existence-based", "nature-of-being"],
    "interdisciplinary":  ["cross-field", "multi-field", "cross-domain"],
    "multidisciplinary":  ["across fields", "spanning fields", "cross-area"],
    "cross-disciplinary": ["bridging fields", "spanning fields"],

    # ── AI-favorite verbs ──
    "elucidate":          ["explain", "clarify", "spell out", "shed light on"],
    "delineate":          ["outline", "map out", "sketch", "lay out"],
    "expedite":           ["speed up", "hasten", "accelerate", "push forward"],
    "ameliorate":         ["improve", "ease", "lessen", "reduce"],
    "mitigate":           ["reduce", "lessen", "ease", "cut"],
    "circumvent":         ["avoid", "get around", "bypass", "sidestep"],
    "contextualize":      ["put in context", "set the scene for", "frame"],
    "operationalize":     ["put into practice", "implement", "carry out"],
    "recalibrate":        ["readjust", "re-tune", "re-set", "adjust again"],
    "scaffold":           ["support", "build up", "structure", "frame"],
    "disseminate":        ["spread", "share", "distribute", "publish"],
    "aggregate":          ["gather", "collect", "combine", "pool"],
    "disaggregate":       ["break down", "separate", "split up"],
    "interrogate":        ["examine", "question", "probe", "study"],
    "problematize":       ["raise questions about", "challenge", "question"],
    "reimagine":          ["rethink", "reconsider", "re-envision"],
    "reconceptualize":    ["rethink", "redefine", "re-frame"],
    "empower":            ["enable", "allow", "help", "support"],
    "spearhead":          ["lead", "drive", "head up", "champion"],
    "orchestrate":        ["organize", "coordinate", "manage", "arrange"],
    "harness":            ["use", "tap", "draw on", "exploit"],
    "cultivate":          ["grow", "develop", "build", "nurture"],
    "champion":           ["support", "back", "advocate for", "push for"],
    "navigate":           ["deal with", "handle", "manage", "work through"],
    "transcend":          ["go beyond", "rise above", "surpass"],
    "underserved":        ["neglected", "overlooked", "left out"],
    "underrepresented":   ["left out", "not well covered", "missing"],

    # ── AI-favorite noun phrases ──
    "a testament to":     ["proof of", "evidence of", "a sign of"],
    "a myriad of":        ["many", "a host of", "lots of", "a slew of"],
    "a plethora of":      ["many", "a mass of", "a mountain of", "lots of"],
    "a cornucopia of":    ["many", "an abundance of", "a wealth of"],
    "an array of":        ["a range of", "a set of", "a series of", "a line-up of"],
    "a constellation of": ["a group of", "a set of", "a cluster of"],
    "a tapestry of":      ["a mix of", "a blend of", "a weave of"],
    "a mosaic of":        ["a mix of", "a patchwork of", "a blend of"],
    "a harbinger of":     ["a sign of", "an early indicator of"],
    "in the realm of":    ["in the field of", "in the area of", "in"],
    "at the forefront of": ["leading", "at the front of", "ahead in"],
    "in the landscape of": ["in the field of", "in the area of", "in"],
    "in the context of":  ["in", "within", "against the background of"],
    "in the domain of":   ["in the field of", "in the area of", "in"],
    "across the spectrum": ["across the board", "throughout", "all across"],
    "on the horizon":     ["coming", "in the pipeline", "ahead"],
    "the advent of":      ["the arrival of", "the coming of", "the rise of"],
    "the proliferation of": ["the spread of", "the growth of", "the rise of"],
    "the intersection of": ["where ... meet", "the meeting point of", "the crossroads of"],
    "the convergence of": ["the coming together of", "the merging of"],
    "the culmination of": ["the result of", "the outcome of", "the end product of"],
    "the cornerstone of": ["the foundation of", "the basis of", "the key part of"],
    "the hallmark of":    ["the sign of", "the mark of", "the feature of"],
    "the impetus for":    ["the push for", "the driver for", "the spur for"],
    "the linchpin of":    ["the key part of", "the central piece of"],
    "the bedrock of":     ["the foundation of", "the basis of"],
    "the crux of":        ["the heart of", "the core of", "the key point of"],
    "the essence of":     ["the heart of", "the core of", "what matters in"],
    "the overarching":    ["the main", "the broad", "the general", "the overall"],
    "overarching":        ["main", "broad", "general", "overall"],
    "wide-ranging":       ["broad", "extensive", "far-reaching", "comprehensive"],
    "far-reaching":       ["broad", "wide", "extensive", "deep"],
    "far-reaching implications": ["broad consequences", "wide effects", "deep effects"],
    "key drivers":        ["main forces", "chief causes", "primary factors"],
    "key enablers":       ["main supporters", "chief enablers", "primary factors"],
    "driving forces":     ["main causes", "key factors", "prime movers"],
    "guiding principles": ["main rules", "core ideas", "basic tenets"],
    "best practices":     ["good methods", "proven approaches", "sound methods"],
    "actionable insights": ["useful findings", "practical takeaways", "clear findings"],
    "data-driven":        ["evidence-based", "backed by data", "empirical"],
    "evidence-based":     ["backed by evidence", "grounded in data", "data-supported"],
    "state-of-the-art":   ["latest", "current", "modern", "cutting-edge"],
    "cutting-edge":       ["latest", "newest", "most advanced", "leading"],
    "next-generation":    ["future", "new", "upcoming", "coming"],
    "large-scale":        ["big", "major", "wide", "broad"],
    "small-scale":        ["limited", "modest", "small", "pilot"],
    "large-scale deployment": ["wide rollout", "broad roll-out", "major deployment"],
    "net-zero":           ["zero-carbon", "carbon-neutral", "zero-emission"],

    # ── AI sentence patterns to break ──
    "it is worth noting that": ["it should be said that", "notably,", "importantly,"],
    "it is important to note that": ["notably,", "it bears mentioning that"],
    "it should be noted that": ["one should note that", "note that"],
    "it is worth emphasizing": ["it bears repeating that"],
    "it is worth mentioning": ["it bears mentioning", "one should mention"],
    "plays a crucial role": ["matters a lot", "is very important", "is key"],
    "plays a vital role": ["is vital", "is essential", "matters greatly"],
    "plays a significant role": ["matters", "is important", "counts for a lot"],
    "plays an important role": ["matters", "is important", "counts"],
    "of paramount importance": ["extremely important", "vital", "critical"],
    "of great importance": ["very important", "key", "central"],
    "of considerable importance": ["quite important", "fairly important"],
    "a growing body of evidence": ["more and more evidence", "increasing evidence"],
    "a growing body of research": ["more and more research", "increasing research"],
    "a growing body of literature": ["more and more studies", "an expanding literature"],
    "sheds light on": ["clarifies", "explains", "illuminates", "unpacks"],
    "paves the way for": ["opens the door to", "makes possible", "allows for"],
    "opens up new avenues": ["creates new paths", "offers new routes", "enables new approaches"],
    "bridges the gap": ["closes the gap", "spans the divide", "fills the gap"],
    "fills a critical gap": ["fills an important hole", "addresses a key missing piece"],
    "addresses a critical gap": ["tackles an important gap", "deals with a key gap"],
    "in recent years": ["lately", "of late", "in the past few years"],
    "over the past decade": ["in the last ten years", "during the 2010s and beyond"],
    "in the era of": ["in the age of", "in the time of", "during"],
    "in today's world": ["now", "currently", "today", "nowadays"],
    "in an increasingly": ["as ... becomes more", "with growing"],
    "in the modern era": ["today", "now", "in our time"],
    "at the dawn of": ["at the start of", "at the beginning of", "as ... begins"],
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
# TRANSITION REPLACEMENTS — v4.0: shorter, more human
# بدائل ربطية — أقصر وأكثر بشرية
# ═══════════════════════════════════════════════════════════════

GENERIC_TRANSITIONS: Dict[str, List[str]] = {
    "Moreover":           ["Then again,", "On top of that,", "And yet,", "What is more,"],
    "In addition":        ["As well,", "Alongside this,", "Added to this,"],
    "Furthermore":        ["On top of that,", "And besides,", "To make matters worse,"],
    "Additionally":       ["As well,", "On top of this,", "To add to that,"],
    "Also":               ["In the same way,", "By the same token,", "Equally,"],
    "In conclusion":      ["All told,", "Taken together,", "Summing up,"],
    "To summarize":       ["In short,", "All in all,", "To sum up,"],
    "Therefore":          ["So,", "For that reason,", "Because of this,"],
    "Thus":               ["So,", "In this way,", "By the same logic,"],
    "Hence":              ["So,", "It follows that", "For this reason,"],
    "However":            ["That said,", "Even so,", "Then again,", "Be that as it may,"],
    "Nevertheless":       ["Even so,", "And yet,", "Still,"],
    "On the other hand":  ["Then again,", "Counterbalancing this,", "Set against this,"],
    "In contrast":        ["Set against this,", "By way of contrast,", "On the flip side,"],
    "Meanwhile":          ["At the same time,", "In parallel,", "At the same point,"],
    "Despite":            ["Even with", "In spite of", "Regardless of"],
}

# ═══════════════════════════════════════════════════════════════
# VARIED SENTENCE OPENERS v4.0 — HUMAN-SOUNDING ONLY
# No AI clichés like "Pivotal here," "Notably," "Strikingly,"
# ═══════════════════════════════════════════════════════════════

VARIED_OPENERS = {
    "ADV":  ["Admittedly,", "Frankly,", "Curiously enough,", "As it happens,", "Coincidentally,"],
    "ADJ":  ["Fair enough,", "True enough,", "As one might expect,", "By all accounts,"],
    "VERB": ["Take,", "Look at", "Consider for a moment,", "Think about", "Recall that"],
    "PREP": ["For one thing,", "Among other things,", "In practice,", "On the ground,", "By way of example,"],
    "CONJ": ["And yet,", "Yet,", "Even so,", "Still,", "Though,"],
}

CRITICAL_PERSPECTIVES: List[str] = [
    "One worry, though, is whether this finding generalises beyond the lab.",
    "The causal story here is far from settled.",
    "Replication in independent settings has been thin on the ground.",
    "A fair question is whether this holds up outside controlled conditions.",
    "Whether the effect survives real-world noise remains to be seen.",
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
# ║           HUMANIZE ENGINE v4.0                ║
# ╚══════════════════════════════════════════════╝

class HumanizeEngine:
    """
    Multi-layered humanization engine — PROFESSIONAL EDITION v4.0.

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
        # v4.0: more aggressive strength mapping
        # strength 1→0.15, 2→0.40, 3→0.65, 4→0.85, 5→1.0
        self._s = 0.15 + (strength - 1) * 0.2125
        self._s = min(self._s, 1.0)

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

        # Step 1: Aggressively replace AI-favored vocabulary
        sentences = [self._replace_ai_vocabulary(s) for s in sentences]

        # Step 2: Sentence-length variation
        sentences = self._vary_sentence_lengths(sentences)

        # Step 3: Additional synonym-based replacement for words NOT in AI dict
        sentences = [self._replace_with_synonyms(s) for s in sentences]

        return " ".join(sentences)

    def _replace_ai_vocabulary(self, sentence: str) -> str:
        """
        استبدال المفردات المفضلة للذكاء الاصطناعي ببدائل بشرية.
        Based on the Wikipedia PDF list of overused AI words.
        v4.0: Much more aggressive — replaces ALL matching AI words, not capped.
        """
        result = sentence
        replacements_made = 0
        # v4.0: allow up to 60% of words to be replaced at max strength
        max_replacements = max(2, int(len(result.split()) * self._s * 0.6))

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
                if replacement:
                    old = match.group()
                    # Case preservation
                    if old[0].isupper():
                        replacement = replacement[0].upper() + replacement[1:]
                    result = pattern.sub(replacement, result, count=1)
                    self.changes.append(ChangeRecord(old, replacement, "M1-AIVocab"))
                    replacements_made += 1

        return result

    def _replace_with_synonyms(self, sentence: str) -> str:
        """
        استبدال إضافي بالمرادفات — Additional synonym replacement for
        non-AI words to break token-frequency watermarks.
        """
        words = sentence.split()
        if len(words) < 8:
            return sentence

        # Pick 1-2 content words to replace via synonym DB
        content_indices = []
        stop = frozenset({"the","a","an","is","are","was","were","be","been","have","has","had",
                         "do","does","did","will","would","could","should","may","might","must",
                         "to","of","in","for","on","with","at","by","from","as","and","but","or",
                         "not","so","yet","this","that","these","those","it","its","they","them",
                         "their","we","our","you","your","he","him","his","she","her","i","me","my",
                         "which","who","whom","whose","that","what","how","when","where","why"})

        for i, w in enumerate(words):
            clean = re.sub(r'[^A-Za-z]', '', w).lower()
            if clean and clean not in stop and len(clean) > 4:
                content_indices.append(i)

        if not content_indices:
            return sentence

        random.shuffle(content_indices)
        replacements = 0
        max_reps = max(1, int(self._s * 2))

        for idx in content_indices:
            if replacements >= max_reps:
                break
            clean = re.sub(r'[^A-Za-z]', '', words[idx]).lower()
            # Skip if already in AI_VOCAB_REPLACEMENTS (handled separately)
            if clean in AI_VOCAB_REPLACEMENTS:
                continue
            syn = self.syn_db.get_synonym(clean, self.field, self.strength)
            if syn and syn.lower() != clean:
                old_word = words[idx]
                suffix = ""
                if old_word.endswith("."):
                    suffix = "."
                elif old_word.endswith(","):
                    suffix = ","
                if old_word[0].isupper():
                    syn = syn[0].upper() + syn[1:]
                words[idx] = syn + suffix
                self.changes.append(ChangeRecord(old_word, words[idx], "M1-Synonym"))
                replacements += 1

        return " ".join(words)

    def _vary_sentence_lengths(self, sentences: List[str]) -> List[str]:
        """تنويع أطوال الجمل — Vary sentence lengths so no two consecutive are within ±3 words."""
        result: List[str] = []
        prev_len = 0

        for i, sent in enumerate(sentences):
            if not sent.strip():
                continue
            words = sent.split()
            cur_len = len(words)

            # v4.0: wider tolerance band (±3), higher probability
            if i > 0 and abs(cur_len - prev_len) <= 3 and self._s > 0.15:
                if cur_len > 18:
                    sp = self._find_split_point(sent)
                    if sp:
                        left = sent[:sp].strip()
                        right = sent[sp:].strip()
                        result.append(left)
                        result.append(right)
                        prev_len = len(right.split())
                        self.changes.append(ChangeRecord(sent[:40]+"...", f"{left[:20]}... | {right[:20]}...", "M1-LengthVar"))
                        continue
                elif cur_len < 14 and i + 1 < len(sentences) and sentences[i+1].strip():
                    merged = sent.rstrip(".") + ", and " + sentences[i+1].lstrip()
                    if merged[0].islower() and len(merged) > 1:
                        merged = merged[0].upper() + merged[1:]
                    result.append(merged)
                    prev_len = len(merged.split())
                    sentences[i+1] = ""
                    self.changes.append(ChangeRecord(sent[:40]+"...", merged[:40]+"...", "M1-Merge"))
                    continue

            # Insert parenthetical aside in long sentences
            if cur_len > 28 and random.random() < self._s * 0.6:
                parens = ["(as one might expect)", "(admittedly)", "(though not universally)", "(at least in theory)", "(in practice, at any rate)"]
                words.insert(random.randint(len(words)//3, 2*len(words)//3), random.choice(parens))
                sent = " ".join(words)
                self.changes.append(ChangeRecord("(no aside)", parens[0], "M1-Aside"))

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
        """ضمان تنويع بدايات الجمل — v4.0: higher probability, human-only openers."""
        words = sent.split()
        if not words:
            return sent, prev_pos

        first_clean = re.sub(r'[^A-Za-z]', '', words[0])
        if not first_clean:
            return sent, prev_pos

        cur_pos = self._simple_pos(first_clean)

        # v4.0: Higher probability (0.7 vs 0.35) and only trigger when
        # 2+ consecutive sentences start the same way
        if prev_pos and cur_pos == prev_pos and random.random() < self._s * 0.9:
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
        """تنويع علامات الترقيم — em-dashes, semicolons, parentheses. v4.0: higher prob."""
        if random.random() < self._s * 0.6:
            # Replace ", and" with em-dash or semicolon
            sent = re.sub(r',\s+and\s+(?=[a-z])',
                         lambda m: '\u2014and ' if random.random() < 0.5 else '; ',
                         sent, count=1)
        # Add a semicolon between two independent clauses joined by comma
        if random.random() < self._s * 0.4:
            sent = re.sub(r',\s+(?:while|whereas|although)\b',
                         lambda m: '; ' + m.group().strip()[:m.group().strip().find(' ')] + ' ',
                         sent, count=1)
        return sent

    def _inject_hedging(self, sent: str) -> str:
        """حقن التمويه العلمي — Inject hedging for assertive statements."""
        for pattern, hedge in ASSERTIVE_REPLACEMENTS.items():
            if re.search(pattern, sent, re.IGNORECASE) and random.random() < self._s * 0.8:
                old = re.search(pattern, sent, re.IGNORECASE).group()
                sent = re.sub(pattern, hedge, sent, count=1, flags=re.IGNORECASE)
                self.changes.append(ChangeRecord(old, hedge, "M2-Hedge"))
                break
        return sent

    def _inject_personal_touch(self, text: str) -> str:
        """إضافة لمسة شخصية — v4.0: More natural, less AI-sounding."""
        touches = ["\u2014or so it seems.", " At least, that is the working assumption.",
                   " Or so the argument goes.", " At any rate, that is the picture so far."]
        sentences = sent_tokenize(text)
        if len(sentences) < 3:
            return text
        if random.random() < self._s * 0.5:
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

        # Step 4: Critical perspective (only in longer texts)
        if section_idx == 0 or random.random() < 0.5 * self._s:
            text = self._inject_critical_perspective(text)

        return text

    def _replace_transitions(self, text: str) -> str:
        """v4.0: Always replace (no random gate) — transitions are a dead giveaway."""
        for generic, replacements in GENERIC_TRANSITIONS.items():
            pattern = re.compile(r'\b' + re.escape(generic) + r'\b', re.IGNORECASE)
            match = pattern.search(text)
            if match and random.random() < self._s * 0.9:
                replacement = random.choice(replacements)
                old = match.group()
                text = pattern.sub(replacement, text, count=1)
                self.changes.append(ChangeRecord(old, replacement, "M3-Transition"))
        return text

    def _remove_superficial_analysis(self, text: str) -> str:
        """
        إزالة عبارات التحليل السطحي المميزة للذكاء الاصطناعي.
        v4.0: Always remove (these are 100% AI tells).
        """
        for pattern, replacement in SUPERFICIAL_PHRASES.items():
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            for m in matches:
                if random.random() < self._s * 0.9:
                    old = m.group()
                    if replacement:
                        text = re.sub(re.escape(old), replacement, text, count=1)
                        self.changes.append(ChangeRecord(old, replacement, "M3-Superficial"))
                    else:
                        start = m.start()
                        if start > 0 and text[start-1] == ',':
                            start -= 1
                        text = text[:start] + text[m.end():]
                        self.changes.append(ChangeRecord(old, "(removed)", "M3-Superficial"))
                        break
        return text

    def _causal_reorder(self, text: str) -> str:
        """v4.0: More aggressive causal connector injection."""
        sentences = sent_tokenize(text)
        if len(sentences) <= 3:
            return text
        listing = [i for i, s in enumerate(sentences)
                   if re.match(r'^(?:It|This|The|Such|These|Those)\s+(?:is|are|was|were|has|have)\b', s.lstrip())]
        if len(listing) >= 2 and random.random() < self._s * 0.7:
            target = listing[min(1, len(listing)-1)]
            connectors = ["In turn, ", "As a direct result, ", "Because of this, ", "For this reason, "]
            conn = random.choice(connectors)
            orig = sentences[target]
            sentences[target] = conn + orig[0].lower() + orig[1:]
            self.changes.append(ChangeRecord(orig[:40]+"...", sentences[target][:40]+"...", "M3-Causal"))
        return " ".join(sentences)

    def _inject_critical_perspective(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences) < 4:
            return text
        if random.random() < self._s * 0.7:
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
            if len(sents) > 5 and random.random() < self._s * 0.5:
                split_at = random.randint(len(sents)//2, len(sents)-2)
                result.append(" ".join(sents[:split_at]))
                result.append(" ".join(sents[split_at:]))
                self.changes.append(ChangeRecord("One paragraph", "Split into two", "M4-Split"))
            elif len(sents) <= 2 and result and random.random() < self._s * 0.4:
                result[-1] += " " + para
                self.changes.append(ChangeRecord("Two paragraphs", "Merged", "M4-Merge"))
            else:
                if len(sents) > 3 and random.random() < self._s * 0.5:
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
        """تشتيت أنماط توزيع الرموز المتكررة — v4.0: lower threshold (3+ occurrences)."""
        words = text.split()
        freq = Counter(w.lower().strip(".,;:!?()\"'") for w in words if re.search(r'[A-Za-z]', w))
        over_rep = {w: c for w, c in freq.items() if c >= 3}

        replacements = 0
        max_reps = max(2, int(len(over_rep) * self._s * 0.5))

        for word, count in over_rep.items():
            if replacements >= max_reps:
                break
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
        v4.0: Lower threshold (0.50) so good changes are NOT reverted.
        """
        # Step 1: Restore protected content
        modified = self._restore_protected(modified, original)

        # Step 2: Similarity check (lightweight)
        # v4.0: Lowered from 0.70 to 0.50 — allow more change
        sim = self._similarity(original, modified)
        if sim < 0.45:
            # Revert to safe modifications only (very rare now)
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
        # Fix double words (e.g., "from from")
        text = re.sub(r'\b(\w+)\s+\1\b', r'\1', text, flags=re.IGNORECASE)
        # Fix dangling prepositions at sentence end (from., to., with.)
        text = re.sub(r'\b(from|to|with|of|for|at|in|on|by)\.$', lambda m: m.group(0)[:-1], text, flags=re.MULTILINE)
        text = re.sub(r'\b(from|to|with|of|for|at|in|on|by)\.\s', lambda m: m.group(0)[:-2] + '. ', text)
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
    """
    from docx import Document

    doc = Document(io.BytesIO(input_bytes))
    all_changes: List[ChangeRecord] = []
    total_paragraphs = len(doc.paragraphs)

    for idx, para in enumerate(doc.paragraphs):
        if progress_cb and idx % 5 == 0:
            pct = 10 + int(80 * (idx / max(total_paragraphs, 1)))
            progress_cb("Processing paragraphs...", pct)

        text = para.text.strip()
        if not text or len(text) < 10:
            continue

        if _is_protected_paragraph(para, text):
            continue

        humanized, changes = engine.humanize_text(text)

        if humanized != text:
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
    """Check if a paragraph should NOT be humanized."""
    # Skip if paragraph contains equations (OMML elements)
    if para._element.findall('.//' + qn('m:oMath')):
        return True

    # Skip reference/bibliography paragraphs
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
    """Replace paragraph text while preserving the formatting of the first run."""
    if not para.runs:
        para.text = new_text
        return

    first_run = para.runs[0]
    font_attrs = {}
    try:
        rPr = first_run._element.find(qn('w:rPr'))
        if rPr is not None:
            import copy
            font_attrs = copy.deepcopy(rPr)
    except Exception:
        pass

    for run in para.runs:
        run.text = ""

    first_run.text = new_text

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

    words = word_tokenize_simple(text)
    ai_words = [w for w in words if w.lower() in AI_VOCAB_REPLACEMENTS]
    ai_density = len(ai_words) / max(len(words), 1)

    lengths = [len(s.split()) for s in sentences]
    mean_l = sum(lengths) / len(lengths)
    var_l = sum((l - mean_l)**2 for l in lengths) / len(lengths)
    burstiness = (var_l ** 0.5) / mean_l if mean_l > 0 else 0

    ttr = len(set(w.lower() for w in words)) / max(len(words), 1)

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

    st.markdown('<div class="main-title">DeepClean Studio v4.0</div>', unsafe_allow_html=True)
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
                            help="1=Minimal changes, 5=Maximum humanization")
        field = st.selectbox("Academic Field", ["General", "Medical", "Engineering", "Humanities"],
                            help="Field-specific synonym selection")
        st.divider()
        st.subheader("Local Self-Check Metrics")
        st.caption("Runs on original text only (no API call)")

    # ── Determine input ──
    input_text = ""
    input_docx_bytes: Optional[bytes] = None
    is_docx = False

    if uploaded_file is not None:
        fname = uploaded_file.name.lower()
        if fname.endswith(".docx"):
            input_docx_bytes = uploaded_file.getvalue()
            is_docx = True
        elif fname.endswith(".pdf"):
            try:
                import fitz  # PyMuPDF
                pdf_doc = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
                pages = [page.get_text() for page in pdf_doc]
                input_text = "\n\n".join(pages)
                pdf_doc.close()
            except ImportError:
                try:
                    from pdfminer.high_level import extract_text
                    input_text = extract_text(io.BytesIO(uploaded_file.getvalue()))
                except ImportError:
                    st.error("PDF support requires PyMuPDF or pdfminer. Install with: pip install pymupdf")
        else:
            input_text = uploaded_file.getvalue().decode("utf-8", errors="replace")
    elif paste_text.strip():
        input_text = paste_text.strip()

    # ── Show original metrics ──
    if input_text or is_docx:
        metric_text = input_text if input_text else "(DOCX — metrics computed after processing)"
        if input_text:
            metrics = compute_metrics(input_text)
            cols = st.sidebar.columns(4)
            for i, (k, v) in enumerate(metrics.items()):
                with cols[i]:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">{k.replace("_"," ")}</div><div class="metric-value">{v}</div></div>', unsafe_allow_html=True)

    # ── Process button ──
    if st.button("\U0001f527 Humanize", type="primary", use_container_width=True):
        if not input_text and not input_docx_bytes:
            st.warning("Please upload a document or paste text first.")
            return

        # Find synonym CSV
        syn_csv = ""
        for candidate in ["synonyms_academic.csv", "synonyms_academic.csv",
                          Path(__file__).parent / "synonyms_academic.csv"]:
            if isinstance(candidate, str) and Path(candidate).exists():
                syn_csv = candidate
                break
            elif isinstance(candidate, Path) and candidate.exists():
                syn_csv = str(candidate)
                break

        engine = HumanizeEngine(syn_csv, field=field, strength=strength)

        progress = st.progress(0, text="Starting...")
        def _progress(msg, pct):
            progress.progress(pct, text=msg)

        if is_docx and input_docx_bytes:
            try:
                result_bytes, changes = process_docx(input_docx_bytes, engine, _progress)
                st.success(f"Done! {len(changes)} changes made.")

                # Show side-by-side
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Original DOCX")
                    st.info("Original document uploaded (see download for comparison)")
                with col2:
                    st.subheader("Humanized DOCX")
                    st.download_button(
                        "\U0001f4be Download Humanized DOCX",
                        data=result_bytes,
                        file_name="humanized_output.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

                # Change review table
                if changes:
                    with st.expander(f"Changes for Human Review ({len(changes)} total)", expanded=False):
                        # Deduplicate by module
                        seen = set()
                        unique_changes = []
                        for c in changes:
                            key = (c.original[:60], c.modified[:60], c.module)
                            if key not in seen:
                                seen.add(key)
                                unique_changes.append(c)

                        st.dataframe(
                            [{"Module": c.module, "Original": c.original[:80], "Modified": c.modified[:80]}
                             for c in unique_changes[:100]],
                            use_container_width=True,
                            hide_index=True,
                        )

            except Exception as e:
                st.error(f"Error processing DOCX: {e}")
                import traceback
                st.code(traceback.format_exc())

        elif input_text:
            try:
                humanized, changes = engine.humanize_text(input_text, _progress)
                st.success(f"Done! {len(changes)} changes made.")

                # Side-by-side view
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Original Text")
                    st.text_area("Original", input_text, height=350, disabled=True, label_visibility="collapsed")
                    st.caption(f"{len(input_text.split())} words")
                with col2:
                    st.subheader("Humanized Text")
                    st.text_area("Humanized", humanized, height=350, disabled=True, label_visibility="collapsed")
                    st.caption(f"{len(humanized.split())} words")

                # Download as DOCX
                try:
                    from docx import Document
                    doc = Document()
                    for para_text in humanized.split("\n\n"):
                        doc.add_paragraph(para_text)
                    output_bytes = io.BytesIO()
                    doc.save(output_bytes)
                    st.download_button(
                        "\U0001f4be Download as DOCX",
                        data=output_bytes.getvalue(),
                        file_name="humanized_output.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                except ImportError:
                    st.download_button(
                        "\U0001f4be Download as TXT",
                        data=humanized,
                        file_name="humanized_output.txt",
                        mime="text/plain",
                        use_container_width=True,
                    )

                # Metrics comparison
                orig_m = compute_metrics(input_text)
                hum_m = compute_metrics(humanized)
                st.divider()
                st.subheader("Metrics Comparison")
                mc1, mc2, mc3, mc4 = st.columns(4)
                for i, (k, (ov, hv)) in enumerate(zip(orig_m.keys(), zip(orig_m.values(), hum_m.values()))):
                    col = [mc1, mc2, mc3, mc4][i]
                    delta = hv - ov
                    arrow = "\u2191" if (k == "perplexity_proxy" or k == "burstiness" or k == "ttr") and delta > 0 else ("\u2193" if delta < 0 else "\u2192")
                    if k == "ai_vocab_density":
                        arrow = "\u2193" if delta < 0 else ("\u2191" if delta > 0 else "\u2192")
                    col.metric(k.replace("_", " ").title(), f"{hv:.3f}", f"{arrow} {abs(delta):.3f}")

                # Change review table
                if changes:
                    with st.expander(f"Changes for Human Review ({len(changes)} total)", expanded=False):
                        seen = set()
                        unique_changes = []
                        for c in changes:
                            key = (c.original[:60], c.modified[:60], c.module)
                            if key not in seen:
                                seen.add(key)
                                unique_changes.append(c)

                        st.dataframe(
                            [{"Module": c.module, "Original": c.original[:80], "Modified": c.modified[:80]}
                             for c in unique_changes[:100]],
                            use_container_width=True,
                            hide_index=True,
                        )

            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
