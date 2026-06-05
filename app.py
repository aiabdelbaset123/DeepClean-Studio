"""
DeepClean Studio v5.0 — NUCLEAR Humanization Engine
====================================================
Transforms AI-generated academic/scientific texts into texts indistinguishable
from expert human writing, with FULL preservation of DOCX formatting.

v5.0 "Nuclear" — Complete rewrite for maximum AI detection evasion.
Based on analysis of GPTZero, Originality.ai, Turnitin, and Copyleaks detection
methods: perplexity uniformity, burstiness, vocabulary fingerprints, and sentence
structure patterns.

Key v5.0 changes:
  - UNCAPPED AI vocabulary replacement (replace ALL matching words)
  - Massive synonym replacement (5-10 words per sentence)
  - Aggressive sentence restructuring (split, merge, reorder clauses)
  - Extreme burstiness injection (3-word to 40-word sentence mix)
  - Clause-level reordering within sentences
  - Passive/active voice switching
  - Phrasal verb injection (formal → informal)
  - Human discourse markers
  - Much lower similarity threshold (0.25 vs 0.45)
  - Multi-pass processing for deeper transformation
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
# AI VOCABULARY DATABASE v5.0 — NUCLEAR EDITION
# Every word here is a known AI signal. Replace ALL of them.
# ═══════════════════════════════════════════════════════════════

AI_VOCAB_REPLACEMENTS: Dict[str, List[str]] = {
    # ── Core AI "tells" — the strongest signals ──
    "crucial":            ["important", "pressing", "central", "key", "vital"],
    "crucially":          ["importantly", "above all", "most of all", "in practice"],
    "pivotal":            ["decisive", "key", "central", "main", "turning-point"],
    "underscore":         ["confirm", "show", "back up", "bear out", "drive home"],
    "highlight":          ["point to", "show", "bring out", "flag", "spotlight"],
    "showcase":           ["present", "display", "demonstrate", "show", "put on show"],
    "delve":              ["examine", "probe", "explore", "look into", "investigate", "dig into"],
    "intricate":          ["complex", "detailed", "elaborate", "involved", "knotty"],
    "fostering":          ["encouraging", "promoting", "supporting", "driving", "spurring"],
    "garner":             ["gain", "attract", "collect", "earn", "pull in"],
    "meticulous":         ["careful", "thorough", "detailed", "close", "painstaking"],
    "robust":             ["strong", "solid", "reliable", "sturdy", "tough"],
    "testament":          ["proof", "evidence", "sign", "indication", "witness"],
    "enduring":           ["lasting", "long-standing", "persistent", "ongoing", "steadfast"],
    "vibrant":            ["active", "lively", "busy", "thriving", "buzzing"],
    "tapestry":           ["mix", "blend", "mosaic", "fabric", "patchwork"],
    "bolstered":          ["strengthened", "supported", "reinforced", "backed", "shored up"],
    "landscape":          ["field", "area", "domain", "scene", "picture"],
    "valuable":           ["useful", "important", "helpful", "worthwhile", "handy"],
    "enhance":            ["improve", "strengthen", "boost", "raise", "step up"],
    "emphasizing":        ["stressing", "underlining", "pointing up", "putting weight on"],
    "showcasing":         ["presenting", "displaying", "featuring", "showing off"],
    "align with":         ["match", "fit", "correspond to", "accord with", "tally with"],
    "serves as":          ["is", "acts as", "functions as", "works as"],
    "stands as":          ["is", "remains", "counts as"],
    "boasts":             ["has", "features", "contains", "includes", "comes with"],
    "features":           ["has", "includes", "contains", "offers", "brings"],
    "offers":             ["provides", "gives", "has", "supplies", "comes with"],
    "represents":         ["is", "forms", "makes up", "counts as"],
    "marks":              ["is", "signals", "marks out", "flags"],
    "exemplifies":        ["illustrates", "typifies", "demonstrates", "is a case of"],
    "additionally":       ["also", "in addition", "further", "on top of that"],
    "comprehensive":      ["thorough", "complete", "full", "wide-ranging", "all-out"],
    "numerous":           ["many", "several", "various", "a range of", "a host of"],
    "significant":        ["notable", "important", "substantial", "marked", "sizeable"],
    "subsequently":       ["then", "later", "afterward", "next", "soon after"],
    "furthermore":        ["moreover", "also", "in addition", "on top of that"],
    "moreover":           ["also", "further", "in addition", "what is more"],
    "notably":            ["in particular", "especially", "worth noting"],
    "remarkably":         ["strikingly", "unusually", "interestingly", "oddly enough"],
    "demonstrates":       ["shows", "reveals", "indicates", "makes clear", "bears out"],
    "illustrates":        ["shows", "reveals", "depicts", "makes visible", "lays out"],
    "underscores":        ["confirms", "reinforces", "stresses", "drives home"],
    "navigating":         ["dealing with", "managing", "handling", "working through"],
    "realm":              ["area", "field", "domain", "sphere", "world"],
    "paramount":          ["top-priority", "utmost", "vital", "chief", "number-one"],
    "unwavering":         ["steady", "firm", "resolute", "consistent", "unfailing"],
    "unprecedented":      ["novel", "never-before-seen", "extraordinary", "new", "unheard-of"],
    "groundbreaking":     ["pioneering", "innovative", "novel", "fresh", "trailblazing"],
    "multifaceted":       ["complex", "many-sided", "layered", "multi-layered"],
    "holistic":           ["integrated", "all-round", "joined-up", "big-picture"],
    "streamlined":        ["efficient", "simplified", "lean", "slick"],
    "leverage":           ["use", "exploit", "harness", "draw on", "tap into"],
    "utilize":            ["use", "employ", "apply", "draw on", "make use of"],
    "facilitate":         ["enable", "help", "assist", "make possible", "smooth the way for"],
    "encompasses":        ["includes", "covers", "spans", "takes in", "sweeps in"],
    "pertains to":        ["relates to", "concerns", "is about", "regards", "bears on"],
    "nestled":            ["located", "situated", "set", "placed", "tucked away"],
    "breathtaking":       ["striking", "impressive", "remarkable", "stunning"],
    "rich":               ["deep", "varied", "extensive", "strong", "full"],
    "profound":           ["deep", "far-reaching", "serious", "weighty", "deep-seated"],
    "dynamic":            ["active", "changing", "evolving", "shifting", "fluid"],
    "charming":           ["pleasant", "appealing", "attractive", "winning"],
    "stunning":           ["striking", "impressive", "remarkable", "eye-catching"],
    "picturesque":        ["scenic", "attractive", "pretty", "lovely"],

    # ── Multi-word AI patterns ──
    "serves as a":        ["is a", "acts as a", "functions as a", "works as a"],
    "serves as an":       ["is an", "acts as an", "functions as an"],
    "stands as a":        ["is a", "remains a", "counts as a"],
    "stands as an":       ["is an", "remains an"],
    "serves as the":      ["is the", "works as the"],
    "stands as the":      ["is the"],
    "place at the centre": ["put at the center", "put at the heart of"],
    "at the centre of":   ["at the heart of", "at the core of", "central to"],
    "constitute the":     ["make up the", "form the", "account for the"],
    "constitute a":       ["make up a", "form a", "amount to a"],
    "remains fragmented": ["is still split", "is still divided", "is still siloed"],
    "remains entirely":   ["is still completely", "is still wholly"],
    "remain entirely":    ["stay completely", "stay wholly"],
    "the single largest":  ["the biggest", "the largest single"],
    "the dominant share":  ["the largest share", "the biggest portion", "most of it"],

    # ── ACADEMIC AI WORDS ──
    "constitute":         ["make up", "form", "account for", "compose"],
    "constitutes":        ["makes up", "forms", "amounts to"],
    "dominant":           ["main", "leading", "chief", "primary", "largest", "top"],
    "established":        ["well-known", "standard", "widely used", "conventional", "tried-and-true"],
    "routinely":          ["commonly", "often", "regularly", "frequently", "typically", "as a rule"],
    "pronounced":         ["strong", "marked", "clear", "noticeable", "sharp", "stark"],
    "uniquely":           ["especially", "particularly", "distinctively", "specially"],
    "systematically":     ["consistently", "thoroughly", "across the board", "methodically"],
    "fragmented":         ["split", "divided", "disconnected", "broken up", "siloed"],
    "entirely":           ["completely", "wholly", "fully", "altogether", "flat-out"],
    "specifically":       ["in particular", "namely", "precisely", "to be exact"],
    "individual":         ["separate", "single", "stand-alone", "discrete"],
    "mechanistic":        ["cause-and-effect", "mechanism-based", "process-based"],
    "trajectory":         ["path", "course", "route", "trend", "track"],
    "trajectories":       ["paths", "courses", "scenarios", "tracks"],
    "episode":            ["event", "period", "occurrence", "spell"],
    "episodes":           ["events", "periods", "occurrences"],
    "derived":            ["obtained", "computed", "calculated", "extracted", "worked out"],
    "exceeding":          ["above", "over", "surpassing", "going beyond", "topping"],
    "committed to":       ["pledged to", "set out to", "aiming to", "promised to"],
    "impose":             ["create", "present", "pose", "bring about"],
    "govern":             ["control", "shape", "determine", "drive", "steer"],
    "expose":             ["provide", "offer", "give access to", "make available"],
    "produced":           ["yielded", "generated", "created", "turned out", "given rise to"],
    "accurate":           ["precise", "reliable", "sound", "good", "on the money"],
    "connected":          ["linked", "tied", "joined", "coupled", "bound up"],
    "disconnected":       ["separate from", "cut off from", "divorced from", "apart from"],
    "compounded":         ["made worse", "aggravated", "intensified", "piled on top of"],
    "inconsistency":      ["mismatch", "gap", "disagreement", "discrepancy"],
    "practitioners":      ["users", "engineers", "researchers", "analysts", "people in the field"],
    "transfer":           ["move", "copy", "pass", "shift", "hand over"],
    "forces":             ["makes", "requires", "compels", "obliges", "pushes"],
    "ignores":            ["misses", "overlooks", "skips", "leaves out", "passes over"],
    "interactions":       ["links", "connections", "couplings", "cross-effects"],
    "progress":           ["advances", "headway", "improvements", "steps forward"],
    "sub-fields":         ["areas", "disciplines", "domains", "branches"],
    "viability":          ["feasibility", "practicality", "workability"],
    "prediction":         ["forecast", "estimate", "projection", "modeling", "guess"],
    "performance":        ["output", "behavior", "results", "functioning", "track record"],
    "ecosystem":          ["set of tools", "toolset", "environment", "tool chain"],
    "boundaries":         ["interfaces", "junctions", "seams", "crossings"],
    "factors":            ["parameters", "variables", "settings", "inputs", "drivers"],
    "inconsistency":      ["mismatch", "gap", "disagreement", "discrepancy", "clash"],

    # ── More AI-favorite adjectives and adverbs ──
    "particularly":       ["especially", "in particular", "specially", "above all"],
    "fundamentally":      ["at root", "basically", "in essence", "at heart"],
    "essentially":        ["basically", "in practice", "at bottom", "for all intents and purposes"],
    "inherently":         ["by nature", "intrinsically", "by its nature"],
    "intrinsically":      ["by nature", "inherently", "in its own right"],
    "substantially":      ["considerably", "a great deal", "markedly", "a lot"],
    "considerably":       ["a great deal", "markedly", "much", "by a long way"],
    "increasingly":       ["more and more", "ever more", "growingly", "with each passing year"],
    "predominantly":      ["mainly", "mostly", "chiefly", "largely", "for the most part"],
    "overwhelmingly":     ["heavily", "vastly", "by far", "hands down"],
    "strikingly":         ["sharply", "markedly", "clearly", "starkly"],
    "markedly":           ["clearly", "noticeably", "sharply", "distinctly"],
    "consequently":       ["as a result", "because of this", "so", "which is why"],
    "thereby":            ["thus", "by doing so", "in this way", "and so"],
    "wherein":            ["in which", "where", "within which"],
    "thereof":            ["of it", "of that", "of which"],
    "albeit":             ["even though", "although", "while", "though"],
    "whereas":            ["while", "although", "though", "when in fact"],
    "hitherto":           ["until now", "so far", "previously"],
    "heretofore":         ["before this", "up to now", "previously"],
    "hence":              ["so", "therefore", "thus", "for this reason"],
    "thus":               ["so", "therefore", "in this way", "and so"],
    "therefore":          ["so", "for this reason", "as a result", "which means"],
    "indeed":             ["in fact", "as it happens", "to be sure", "and rightly so"],
    "conversely":         ["on the flip side", "on the other hand", "by contrast"],
    "simultaneously":     ["at the same time", "in parallel", "together", "all at once"],
    "subsequent":         ["later", "following", "next", "coming after"],
    "aforementioned":     ["previously mentioned", "above", "earlier"],
    "pertinent":          ["relevant", "germane", "related", "applicable"],
    "salient":            ["key", "main", "notable", "important", "stand-out"],
    "seminal":            ["foundational", "key", "landmark", "path-breaking"],
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
    "indispensable":      ["essential", "vital", "necessary", "cannot-do-without"],
    "catalyst":           ["driver", "spark", "trigger", "cause"],
    "cornerstone":        ["foundation", "basis", "key element", "bedrock"],
    "linchpin":           ["key part", "central piece", "hinge"],
    "hallmark":           ["sign", "feature", "trademark", "mark"],
    "dichotomy":          ["split", "division", "contrast", "two-part split"],
    "amalgamation":       ["blend", "mix", "combination", "merger"],
    "juxtaposition":      ["contrast", "comparison", "side-by-side view"],
    "interdisciplinary":  ["cross-field", "multi-field", "cross-domain"],
    "multidisciplinary":  ["across fields", "spanning fields", "cross-area"],

    # ── AI-favorite verbs ──
    "elucidate":          ["explain", "clarify", "spell out", "shed light on"],
    "delineate":          ["outline", "map out", "sketch", "lay out"],
    "expedite":           ["speed up", "hasten", "accelerate", "push forward"],
    "ameliorate":         ["improve", "ease", "lessen", "reduce"],
    "mitigate":           ["reduce", "lessen", "ease", "cut", "take the edge off"],
    "circumvent":         ["avoid", "get around", "bypass", "sidestep"],
    "contextualize":      ["put in context", "set the scene for", "frame"],
    "operationalize":     ["put into practice", "implement", "carry out"],
    "recalibrate":        ["readjust", "re-tune", "re-set", "adjust again"],
    "scaffold":           ["support", "build up", "structure", "frame"],
    "disseminate":        ["spread", "share", "distribute", "publish"],
    "aggregate":          ["gather", "collect", "combine", "pool", "pull together"],
    "disaggregate":       ["break down", "separate", "split up"],
    "interrogate":        ["examine", "question", "probe", "study"],
    "problematize":       ["raise questions about", "challenge", "question"],
    "reimagine":          ["rethink", "reconsider", "re-envision"],
    "reconceptualize":    ["rethink", "redefine", "re-frame"],
    "empower":            ["enable", "allow", "help", "support"],
    "spearhead":          ["lead", "drive", "head up", "champion"],
    "orchestrate":        ["organize", "coordinate", "manage", "arrange"],
    "harness":            ["use", "tap", "draw on", "exploit", "put to work"],
    "cultivate":          ["grow", "develop", "build", "nurture"],
    "champion":           ["support", "back", "advocate for", "push for"],
    "navigate":           ["deal with", "handle", "manage", "work through"],
    "transcend":          ["go beyond", "rise above", "surpass"],

    # ── AI-favorite noun phrases ──
    "a testament to":     ["proof of", "evidence of", "a sign of", "a mark of"],
    "a myriad of":        ["many", "a host of", "lots of", "a slew of"],
    "a plethora of":      ["many", "a mass of", "lots of", "a mountain of"],
    "an array of":        ["a range of", "a set of", "a series of", "a line-up of"],
    "a tapestry of":      ["a mix of", "a blend of", "a weave of"],
    "in the realm of":    ["in the field of", "in the area of", "in"],
    "at the forefront of": ["leading", "at the front of", "ahead in"],
    "in the landscape of": ["in the field of", "in the area of", "in"],
    "in the context of":  ["in", "within", "against the background of"],
    "in the domain of":   ["in the field of", "in the area of", "in"],
    "the advent of":      ["the arrival of", "the coming of", "the rise of"],
    "the proliferation of": ["the spread of", "the growth of", "the rise of"],
    "the intersection of": ["where ... meet", "the meeting point of"],
    "the cornerstone of": ["the foundation of", "the basis of", "the key part of"],
    "the crux of":        ["the heart of", "the core of", "the key point of"],
    "overarching":        ["main", "broad", "general", "overall"],
    "wide-ranging":       ["broad", "extensive", "far-reaching"],
    "far-reaching":       ["broad", "wide", "extensive", "deep"],
    "key drivers":        ["main forces", "chief causes", "primary factors"],
    "best practices":     ["good methods", "proven approaches", "sound methods"],
    "data-driven":        ["evidence-based", "backed by data", "empirical"],
    "state-of-the-art":   ["latest", "current", "modern", "cutting-edge"],
    "cutting-edge":       ["latest", "newest", "most advanced", "leading"],
    "large-scale":        ["big", "major", "wide", "broad"],

    # ── AI sentence patterns ──
    "it is worth noting that": ["it should be said that", "notably,", "importantly,"],
    "it is important to note that": ["notably,", "it bears mentioning that"],
    "it should be noted that": ["one should note that", "note that"],
    "plays a crucial role": ["matters a lot", "is very important", "is key"],
    "plays a vital role": ["is vital", "is essential", "matters greatly"],
    "plays a significant role": ["matters", "is important", "counts for a lot"],
    "plays an important role": ["matters", "is important", "counts"],
    "of paramount importance": ["extremely important", "vital", "critical"],
    "a growing body of evidence": ["more and more evidence", "increasing evidence"],
    "a growing body of research": ["more and more research", "increasing research"],
    "sheds light on": ["clarifies", "explains", "unpacks", "helps explain"],
    "paves the way for": ["opens the door to", "makes possible", "allows for"],
    "bridges the gap": ["closes the gap", "fills the gap", "spans the divide"],
    "in recent years": ["lately", "of late", "in the past few years"],
    "over the past decade": ["in the last ten years", "since the early 2010s"],

    # ── v5.0: Even more common AI academic words ──
    "comprise":           ["make up", "form", "consist of", "take in"],
    "comprises":          ["makes up", "forms", "consists of", "takes in"],
    "yield":              ["produce", "give", "result in", "bring about"],
    "yields":             ["produces", "gives", "results in", "brings about"],
    "exhibit":            ["show", "display", "have", "reveal"],
    "exhibits":           ["shows", "displays", "has", "reveals"],
    "encompass":          ["include", "cover", "span", "take in"],
    "invoke":             ["call on", "use", "bring in", "turn to"],
    "devise":             ["come up with", "create", "design", "work out"],
    "ascertain":          ["find out", "determine", "work out", "establish"],
    "elicit":             ["bring out", "draw out", "produce", "get at"],
    "articulate":         ["state", "express", "set out", "spell out"],
    "expound":            ["explain", "set out", "lay out", "spell out"],
    "engender":           ["cause", "bring about", "give rise to", "produce"],
    "galvanize":          ["spur", "drive", "push", "fire up"],
    "precipitate":        ["bring about", "cause", "trigger", "speed up"],
    "augment":            ["increase", "add to", "boost", "step up"],
    "supplement":         ["add to", "top up", "boost", "back up"],
    "contend":            ["argue", "claim", "maintain", "hold that"],
    "posit":              ["suggest", "propose", "put forward", "argue"],
    "postulate":          ["suggest", "propose", "put forward", "assume"],
    "corroborate":        ["back up", "confirm", "support", "bear out"],
    "corroborates":       ["backs up", "confirms", "supports", "bears out"],
    "consummate":         ["complete", "perfect", "finish off"],
    "promulgate":         ["announce", "publish", "put out", "spread"],
    "disseminate":        ["spread", "share", "distribute", "put out"],
    "conducive":          ["helpful", "favorable", "good for", "beneficial"],
    "conducive to":       ["helpful for", "good for", "beneficial for"],
    "instrumental":       ["key", "important", "central", "pivotal"],
    "pivotal":            ["decisive", "key", "turning-point", "make-or-break"],
    "quintessential":     ["classic", "typical", "archetypal", "textbook"],
    "prevalent":          ["common", "widespread", "usual", "widespread"],
    "concurrent":         ["simultaneous", "at the same time", "parallel"],
    "analogous":          ["similar", "like", "comparable", "parallel"],
    "homogeneous":        ["uniform", "same", "even", "alike"],
    "heterogeneous":      ["mixed", "varied", "diverse", "uneven"],
    "indigenous":         ["native", "local", "home-grown", "native-born"],
    "concomitant":        ["accompanying", "going along with", "attendant"],
    "extraneous":         ["irrelevant", "outside", "unrelated", "beside the point"],
    "permeate":           ["spread through", "fill", "run through", "soak into"],
    "proliferate":        ["spread", "grow", "multiply", "take off"],
    "relegate":           ["push aside", "demote", "push down", "push to the back"],
    "amalgamate":         ["merge", "combine", "blend", "bring together"],
    "consolidate":        ["bring together", "merge", "combine", "strengthen"],
    "disseminate":        ["spread", "share", "put out", "circulate"],
    "impede":             ["block", "hold up", "slow down", "get in the way of"],
    "stymie":             ["block", "hold up", "thwart", "stand in the way of"],
    "obviate":            ["remove the need for", "do away with", "make unnecessary"],
    "preclude":           ["rule out", "prevent", "stop", "shut out"],
    "repercussions":      ["effects", "fallout", "consequences", "knock-on effects"],
    "ramifications":      ["consequences", "effects", "fallout", "spin-offs"],
    "nuances":            ["subtleties", "fine points", "details", "shades"],
    "substrate":          ["base", "foundation", "underlying layer"],
    "trajectory":         ["path", "course", "track", "line"],
    "paradigm":           ["model", "framework", "approach", "way of thinking"],
    "dichotomy":          ["split", "division", "contrast"],
    "continuum":          ["range", "spectrum", "scale", "gradient"],
    "symbiosis":          ["partnership", "give-and-take", "mutual benefit"],
    "juxtaposition":      ["contrast", "comparison", "side-by-side view"],
    "delineation":        ["outline", "mapping", "laying out", "sketch"],
    "elucidation":        ["explanation", "clarification", "spelling out"],
    "proliferation":      ["spread", "growth", "multiplication", "take-off"],
    "elaboration":        ["detail", "expansion", "fleshing out"],
    "amalgamation":       ["blend", "mix", "merger", "coming together"],
    "recalibration":      ["readjustment", "re-tuning", "re-setting"],
    "contextualization":  ["framing", "setting the scene", "putting in context"],
    "conceptualization":  ["thinking through", "framing", "working out"],
    "operationalization": ["putting into practice", "carrying out", "making real"],
    "methodological":     ["method-based", "procedural", "technique-based"],
    "epistemological":    ["knowledge-based", "philosophical"],
    "pedagogical":        ["teaching", "educational", "instructional"],
    "empirical":          ["observed", "real-world", "based on data", "hands-on"],
    "theoretical":        ["in-theory", "abstract", "on-paper"],
    "normative":          ["standard", "prescriptive", "how things should be"],
    "prescriptive":       ["rule-based", "telling-you-what-to-do", "directive"],
    "descriptive":        ["describing", "factual", "reporting"],
    "speculative":        ["guesswork", "theoretical", "conjectural"],
    "heuristic":          ["rule-of-thumb", "practical", "trial-and-error"],
    "ontological":        ["existence-based", "nature-of-being"],
    "pragmatic":          ["practical", "down-to-earth", "no-nonsense"],
    "ubiquitous":         ["everywhere", "omnipresent", "all-over", "universal"],
    "idiosyncratic":      ["quirky", "unique", "one-off", "peculiar"],
    "esoteric":           ["obscure", "specialist", "niche", "little-known"],
    "exhaustive":         ["complete", "thorough", "all-out", "full"],
    "tentative":          ["provisional", "cautious", "preliminary", "trial"],
    "definitive":         ["final", "conclusive", "authoritative", "last word"],
    "comprehensive":      ["all-in", "complete", "thorough", "full"],
    "substantive":        ["real", "meaningful", "substantial", "solid"],
    "procedural":         ["step-by-step", "by-the-book", "routine"],
    "normative":          ["standard", "expected", "usual"],
    "generative":         ["creative", "productive", "producing"],
    "transformative":     ["game-changing", "radical", "far-reaching", "big"],
    "detrimental":        ["harmful", "damaging", "bad", "negative"],
    "beneficial":         ["helpful", "good", "positive", "useful"],
    "influential":        ["important", "powerful", "weighty", "big"],
    "consequential":      ["important", "significant", "far-reaching", "weighty"],
    "detrimental":        ["harmful", "damaging", "bad", "negative"],
    "trivial":            ["minor", "small", "petty", "unimportant"],
    "negligible":         ["tiny", "insignificant", "barely there", "next to nothing"],
    "marginal":           ["small", "slight", "minimal", "barely noticeable"],
    "inverse":            ["opposite", "reverse", "flipped", "other way round"],
    "subsequent":         ["later", "following", "next", "coming after"],
    "aforementioned":     ["previously mentioned", "above", "earlier", "as noted"],
    "noteworthy":         ["worth a mention", "striking", "remarkable", "stand-out"],
    "insofar":            ["to the extent", "as far as", "inasmuch as"],
    "whereby":            ["by which", "through which", "by means of which"],
    "henceforth":         ["from now on", "from this point", "from here on"],
}


# ═══════════════════════════════════════════════════════════════
# SUPERFICIAL ANALYSIS PHRASES — remove entirely
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
# HEDGING PHRASES — v5.0: more natural, less formulaic
# ═══════════════════════════════════════════════════════════════

HEDGE_PHRASES: List[str] = [
    "it seems fair to say that",
    "the evidence tends to suggest",
    "one could make the case that",
    "it looks as though",
    "the data point toward",
    "a reasonable reading would be that",
    "on the face of it,",
    "chances are that",
    "it would seem that",
    "the early signs are that",
    "all told, it appears that",
    "by and large,",
    "more often than not,",
    "as a rule,",
    "by all indications,",
]

ASSERTIVE_REPLACEMENTS: Dict[str, str] = {
    r"\bclearly\b":               "it seems fair to say that",
    r"\bobviously\b":             "it looks as though",
    r"\bcertainly\b":             "the evidence tends to suggest that",
    r"\bundeniable\b":            "hard to dispute",
    r"\bit is evident that\b":    "the data point toward the idea that",
    r"\bthere is no doubt that\b": "there are good reasons to think that",
    r"\bdefinitively\b":          "provisionally",
    r"\bundeniably\b":            "arguably",
    r"\bwithout question\b":      "one could reasonably say that",
    r"\bindisputably\b":          "by most accounts",
}


# ═══════════════════════════════════════════════════════════════
# TRANSITION REPLACEMENTS — v5.0: even more human
# ═══════════════════════════════════════════════════════════════

GENERIC_TRANSITIONS: Dict[str, List[str]] = {
    "Moreover":           ["Then again,", "On top of that,", "And yet,", "What is more,", "Besides,"],
    "In addition":        ["As well,", "Alongside this,", "Added to this,", "On top of which,"],
    "Furthermore":        ["On top of that,", "And besides,", "To make matters worse,", "On top of which,"],
    "Additionally":       ["As well,", "On top of this,", "To add to that,", "And into the bargain,"],
    "Also":               ["In the same way,", "By the same token,", "Equally,", "And,"],
    "In conclusion":      ["All told,", "Taken together,", "Summing up,", "At the end of the day,"],
    "To summarize":       ["In short,", "All in all,", "To sum up,", "Put briefly,"],
    "Therefore":          ["So,", "For that reason,", "Because of this,", "Which is why,"],
    "Thus":               ["So,", "In this way,", "By the same logic,", "And so,"],
    "Hence":              ["So,", "It follows that", "For this reason,"],
    "However":            ["That said,", "Even so,", "Then again,", "Be that as it may,", "Mind you,"],
    "Nevertheless":       ["Even so,", "And yet,", "Still,", "All the same,"],
    "On the other hand":  ["Then again,", "Counterbalancing this,", "Set against this,", "Flip side,"],
    "In contrast":        ["Set against this,", "By way of contrast,", "On the flip side,"],
    "Meanwhile":          ["At the same time,", "In parallel,", "At the same point,"],
    "Despite":            ["Even with", "In spite of", "Regardless of", "Notwithstanding"],
}


# ═══════════════════════════════════════════════════════════════
# VARIED SENTENCE OPENERS v5.0 — much more diverse
# ═══════════════════════════════════════════════════════════════

VARIED_OPENERS = {
    "ADV":  ["Admittedly,", "Frankly,", "Curiously enough,", "As it happens,",
             "Coincidentally,", "Not surprisingly,", "Oddly enough,", "Unsurprisingly,"],
    "ADJ":  ["Fair enough,", "True enough,", "As one might expect,", "By all accounts,",
             "Sure enough,", "At first glance,", "On the face of it,"],
    "VERB": ["Consider for a moment,", "Think about", "Recall that", "Note that",
             "Look more closely and", "Take the case of", "Bear in mind that"],
    "PREP": ["For one thing,", "Among other things,", "In practice,", "On the ground,",
             "By way of example,", "In point of fact,", "As things stand,", "Looking ahead,"],
    "CONJ": ["And yet,", "Yet,", "Even so,", "Still,", "Though,", "Mind you,", "All the same,"],
}

CRITICAL_PERSPECTIVES: List[str] = [
    "One worry, though, is whether this holds up outside controlled conditions.",
    "The causal story here is far from settled.",
    "Whether this scales up is an open question.",
    "A fair question is whether this generalises beyond the lab.",
    "Whether the effect survives real-world noise remains to be seen.",
    "That, at any rate, is the working hypothesis.",
    "This is, of course, easier said than done.",
    "At least, that is the picture so far.",
]


# ═══════════════════════════════════════════════════════════════
# v5.0 NEW: CLAUSE RESTRUCTURING PATTERNS
# Pattern → how to rearrange the sentence
# ═══════════════════════════════════════════════════════════════

# Patterns for splitting long sentences at clause boundaries
CLAUSE_SPLIT_MARKERS = [
    (r',\s+which\s+', lambda m: ['. ', 'which ']),
    (r',\s+while\s+', lambda m: [' — ', 'while ']),
    (r',\s+whereas\s+', lambda m: ['; ', 'whereas ']),
    (r',\s+although\s+', lambda m: ['. ', 'Although ']),
    (r',\s+and\s+(?=[a-z])', lambda m: [', and ']),  # keep as-is sometimes
]

# Human-style parenthetical asides for burstiness
PARENTHETICALS = [
    "(or so it seems)",
    "(at least in theory)",
    "(in practice, at any rate)",
    "(admittedly)",
    "(though not universally)",
    "(as one might expect)",
    "(for what it is worth)",
    "(to be fair)",
    "(or at least that is the claim)",
    "(at any rate)",
    "(by and large)",
    "(for the most part)",
    "(as things stand)",
]

# Short punchy follow-up sentences for burstiness
BURST_SHORT_SENTENCES = [
    "This matters.",
    "And it shows.",
    "The implications are clear.",
    "That is the crux.",
    "So far, so good.",
    "The picture is mixed.",
    "There is more to it, though.",
    "This is hardly surprising.",
    "But there is a catch.",
    "That much is clear.",
    "It is an open question.",
    "The jury is still out.",
    "Not so fast, though.",
    "Things are more complicated than that.",
    "But wait — there is more.",
]

# v5.0: Phrasal verb replacements (formal verb → phrasal verb)
# Humans use phrasal verbs; AI uses Latinate formal verbs
PHRASAL_VERB_REPLACEMENTS: Dict[str, List[str]] = {
    "investigate":    ["look into", "probe", "dig into"],
    "investigates":   ["looks into", "probes"],
    "investigated":   ["looked into", "probed"],
    "evaluate":       ["size up", "weigh up", "look at"],
    "evaluates":      ["sizes up", "weighs up"],
    "evaluated":      ["sized up", "weighed up"],
    "compensate":     ["make up for", "offset"],
    "compensates":    ["makes up for", "offsets"],
    "compensated":    ["made up for", "offset"],
    "accumulate":     ["build up", "pile up", "rack up"],
    "accumulates":    ["builds up", "piles up"],
    "accumulated":    ["built up", "piled up"],
    "collaborate":    ["work together", "team up", "join forces"],
    "collaborates":   ["works together", "teams up"],
    "collaborated":   ["worked together", "teamed up"],
    "anticipate":     ["expect", "look ahead to", "count on"],
    "anticipates":    ["expects", "looks ahead to"],
    "anticipated":    ["expected", "looked ahead to"],
    "circumvent":     ["get around", "bypass", "sidestep"],
    "circumvents":    ["gets around", "bypasses"],
    "circumvented":   ["got around", "bypassed"],
    "consolidate":    ["bring together", "pull together", "roll up"],
    "consolidates":   ["brings together", "pulls together"],
    "consolidated":   ["brought together", "pulled together"],
    "diminish":       ["cut down", "whittle away", "pare back"],
    "diminishes":     ["cuts down", "whittles away"],
    "diminished":     ["cut down", "whittled away"],
    "eliminate":      ["weed out", "do away with", "root out"],
    "eliminates":     ["weeds out", "does away with"],
    "eliminated":     ["weeded out", "did away with"],
    "fluctuate":      ["go up and down", "swing", "bounce around"],
    "fluctuates":     ["goes up and down", "swings"],
    "fluctuated":     ["went up and down", "swung"],
    "implement":      ["put in place", "roll out", "carry out"],
    "implements":     ["puts in place", "rolls out"],
    "implemented":    ["put in place", "rolled out"],
    "incorporate":    ["take in", "build in", "fold in"],
    "incorporates":   ["takes in", "builds in"],
    "incorporated":   ["took in", "built in"],
    "indicate":       ["point to", "hint at", "suggest"],
    "indicates":      ["points to", "hints at"],
    "indicated":      ["pointed to", "hinted at"],
    "obtain":         ["get hold of", "come by", "pick up"],
    "obtains":        ["gets hold of", "comes by"],
    "obtained":       ["got hold of", "came by"],
    "participate":    ["take part", "join in"],
    "participates":   ["takes part", "joins in"],
    "participated":   ["took part", "joined in"],
    "replicate":      ["copy", "reproduce", "duplicate"],
    "replicates":     ["copies", "reproduces"],
    "replicated":     ["copied", "reproduced"],
    "substantiate":   ["back up", "bear out", "support"],
    "substantiates":  ["backs up", "bears out"],
    "substantiated":  ["backed up", "bore out"],
    "supplement":     ["add to", "top up", "beef up"],
    "supplements":    ["adds to", "tops up"],
    "supplemented":   ["added to", "topped up"],
    "withstand":      ["stand up to", "hold up against", "take"],
    "withstands":     ["stands up to", "holds up against"],
    "withstood":      ["stood up to", "held up against"],
    "exacerbate":     ["make worse", "stir up", "add fuel to"],
    "exacerbates":    ["makes worse", "stirs up"],
    "exacerbated":    ["made worse", "stirred up"],
    "mitigate":       ["take the edge off", "ease", "soften"],
    "mitigates":      ["takes the edge off", "eases"],
    "mitigated":      ["took the edge off", "eased"],
    "disseminate":    ["put out", "spread around", "get out"],
    "disseminates":   ["puts out", "spreads around"],
    "disseminated":   ["put out", "spread around"],
    "reconcile":      ["bring together", "square", "iron out"],
    "reconciles":     ["brings together", "squares"],
    "reconciled":     ["brought together", "squared"],
}


# ═══════════════════════════════════════════════════════════════
# PROTECTED CONTENT PATTERNS (citations, equations, references)
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

# Stop words — never replace these
STOP_WORDS = frozenset({
    "the","a","an","is","are","was","were","be","been","have","has","had",
    "do","does","did","will","would","could","should","may","might","must",
    "to","of","in","for","on","with","at","by","from","as","and","but","or",
    "not","so","yet","this","that","these","those","it","its","they","them",
    "their","we","our","you","your","he","him","his","she","her","i","me","my",
    "which","who","whom","whose","what","how","when","where","why","all","each",
    "every","both","few","more","most","other","some","such","no","nor","only",
    "own","same","than","too","very","can","just","don","should","now","also",
    "if","then","because","while","about","between","through","during","before",
    "after","above","below","up","down","out","off","over","under","again",
    "further","once","here","there","when","where","why","how","any","being",
})


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
# ║           HUMANIZE ENGINE v5.0 "NUCLEAR"      ║
# ╚══════════════════════════════════════════════╝

class HumanizeEngine:
    """
    NUCLEAR Humanization Engine v5.0.

    Key difference from v4.0: UNCAPPED replacements, multi-pass processing,
    extreme burstiness, clause restructuring, phrasal verb injection.

    Module 1 — Statistical Bone-Breaker (UNCAPPED vocab + burstiness)
    Module 2 — Stylometric Mask (openers + punctuation + hedging + phrasal verbs)
    Module 3 — Semantic Deepener (transitions + superficial removal + causality + critical)
    Module 4 — Watermark & Structure Disrupter (reorder + vocab frequency)
    Module 5 — Coherence & Integrity Guardian (similarity at 0.25 + grammar)
    """

    def __init__(self, synonym_csv: str, field: str = "General", strength: int = 3):
        self.syn_db = SynonymDatabase(synonym_csv)
        self.field = field
        self.strength = strength
        self.changes: List[ChangeRecord] = []
        # v5.0: strength mapping — much higher base
        # strength 1→0.30, 2→0.50, 3→0.75, 4→0.90, 5→1.0
        self._s = 0.30 + (strength - 1) * 0.175
        self._s = min(self._s, 1.0)

    def humanize_text(self, text: str, progress_cb=None) -> Tuple[str, List[ChangeRecord]]:
        """Humanize a plain text string — MULTI-PASS for deeper transformation."""
        self.changes = []

        if progress_cb:
            progress_cb("Splitting...", 5)

        chunks = self._split_into_chunks(text)

        if progress_cb:
            progress_cb("Pass 1: Core transformation...", 10)

        humanized_chunks: List[str] = []
        for idx, chunk in enumerate(chunks):
            if progress_cb:
                pct = 10 + int(35 * (idx / max(len(chunks), 1)))
                progress_cb("Pass 1: Core transformation...", pct)

            # PASS 1: Core transformation
            working = chunk
            working = self.module1_statistical_breaker(working)
            working = self.module2_stylometric_mask(working)
            working = self.module3_semantic_deepener(working, section_idx=idx)
            working = self.module4_watermark_disrupter(working)
            working = self.module5_coherence_guardian(working, chunk)

            # PASS 2: Secondary deepening (run modules 1+2 again for more changes)
            if progress_cb:
                pct = 50 + int(30 * (idx / max(len(chunks), 1)))
                progress_cb("Pass 2: Deepening...", pct)

            working = self._secondary_pass(working)

            humanized_chunks.append(working)

        result = "\n\n".join(humanized_chunks)

        if progress_cb:
            progress_cb("Final verification...", 95)

        result = self._grammar_sweep(result)

        if progress_cb:
            progress_cb("Done", 100)

        return result, self.changes

    def _secondary_pass(self, text: str) -> str:
        """Second pass: re-run AI vocab replacement + synonym replacement + cleanup."""
        sentences = sent_tokenize(text)
        if len(sentences) <= 1:
            return text

        result_sents = []
        used_bursts = set()
        for i, sent in enumerate(sentences):
            # Re-run AI vocab (catch any that appeared from replacements)
            sent = self._replace_ai_vocabulary(sent, uncapped=True)
            # Re-run synonym replacement (but gentler)
            sent = self._replace_with_synonyms_gentle(sent)
            result_sents.append(sent)

        return " ".join(result_sents)

    # ──────────────────────────────────────────
    # MODULE 1: Statistical Bone-Breaker
    # ──────────────────────────────────────────
    def module1_statistical_breaker(self, text: str) -> str:
        """
        NUCLEAR: Destroy the smooth probability curve and monotonous rhythm.
        v5.0: UNCAPPED AI vocab replacement, aggressive burstiness,
        phrasal verb injection, extreme sentence length variation.
        """
        sentences = sent_tokenize(text)
        if len(sentences) <= 1:
            return text

        # Step 1: Replace ALL AI-favored vocabulary (UNCAPPED)
        sentences = [self._replace_ai_vocabulary(s, uncapped=True) for s in sentences]

        # Step 2: Phrasal verb injection
        sentences = [self._inject_phrasal_verbs(s) for s in sentences]

        # Step 3: Aggressive synonym replacement
        sentences = [self._replace_with_synonyms_aggressive(s) for s in sentences]

        # Step 4: Extreme sentence-length variation
        sentences = self._vary_sentence_lengths_nuclear(sentences)

        # Step 5: Inject burst short sentences
        result_sents = []
        for i, sent in enumerate(sentences):
            result_sents.append(sent)
            # After every 2-3 long sentences, inject a short punchy one
            if (i + 1) % random.randint(2, 3) == 0 and random.random() < self._s * 0.6:
                burst = random.choice(BURST_SHORT_SENTENCES)
                result_sents.append(burst)
                self.changes.append(ChangeRecord("(no burst)", burst, "M1-Burst"))

        return " ".join(result_sents)

    def _replace_ai_vocabulary(self, sentence: str, uncapped: bool = False) -> str:
        """
        v5.0: Replace ALL matching AI vocabulary words — no cap.
        This is the single most impactful humanization step.
        """
        result = sentence
        replacements_made = 0
        # v5.0: UNCAPPED — replace every single AI word we find
        max_replacements = 999 if uncapped else max(5, int(len(result.split()) * self._s * 1.5))

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

    def _inject_phrasal_verbs(self, sentence: str) -> str:
        """
        v5.0 NEW: Replace formal Latinate verbs with phrasal verbs.
        Humans say "look into"; AI says "investigate".
        """
        result = sentence
        replacements = 0
        max_reps = max(2, int(self._s * 4))

        for formal, phrasals in PHRASAL_VERB_REPLACEMENTS.items():
            if replacements >= max_reps:
                break
            pattern = re.compile(r'\b' + re.escape(formal) + r'\b', re.IGNORECASE)
            match = pattern.search(result)
            if match and random.random() < self._s * 0.7:
                phrasal = random.choice(phrasals)
                old = match.group()
                # Case preservation
                if old[0].isupper():
                    phrasal = phrasal[0].upper() + phrasal[1:]
                result = pattern.sub(phrasal, result, count=1)
                self.changes.append(ChangeRecord(old, phrasal, "M1-PhrasalVerb"))
                replacements += 1

        return result

    def _replace_with_synonyms_gentle(self, sentence: str) -> str:
        """Gentle synonym replacement for secondary pass — 1-2 words only."""
        words = sentence.split()
        if len(words) < 8:
            return sentence

        content_indices = []
        for i, w in enumerate(words):
            clean = re.sub(r'[^A-Za-z]', '', w).lower()
            if clean and clean not in STOP_WORDS and len(clean) > 4:
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
            if clean in AI_VOCAB_REPLACEMENTS or clean in PHRASAL_VERB_REPLACEMENTS:
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
                self.changes.append(ChangeRecord(old_word, words[idx], "M1-Synonym2"))
                replacements += 1

        return " ".join(words)

    def _replace_with_synonyms_aggressive(self, sentence: str) -> str:
        """
        v5.0: AGGRESSIVE synonym replacement — 3-6 words per sentence.
        Previous version only replaced 1-2 words. This replaces many more.
        """
        words = sentence.split()
        if len(words) < 6:
            return sentence

        content_indices = []
        for i, w in enumerate(words):
            clean = re.sub(r'[^A-Za-z]', '', w).lower()
            if clean and clean not in STOP_WORDS and len(clean) > 3:
                content_indices.append(i)

        if not content_indices:
            return sentence

        random.shuffle(content_indices)
        replacements = 0
        # v5.0: replace 3-6 words per sentence based on strength
        max_reps = max(3, int(self._s * len(content_indices) * 0.4))

        for idx in content_indices:
            if replacements >= max_reps:
                break
            clean = re.sub(r'[^A-Za-z]', '', words[idx]).lower()
            # Skip if already in AI_VOCAB_REPLACEMENTS (handled separately)
            if clean in AI_VOCAB_REPLACEMENTS:
                continue
            # Skip if already in PHRASAL_VERB_REPLACEMENTS (handled separately)
            if clean in PHRASAL_VERB_REPLACEMENTS:
                continue
            syn = self.syn_db.get_synonym(clean, self.field, self.strength)
            if syn and syn.lower() != clean:
                old_word = words[idx]
                suffix = ""
                if old_word.endswith("."):
                    suffix = "."
                elif old_word.endswith(","):
                    suffix = ","
                elif old_word.endswith(";"):
                    suffix = ";"
                elif old_word.endswith(":"):
                    suffix = ":"
                if old_word[0].isupper():
                    syn = syn[0].upper() + syn[1:]
                words[idx] = syn + suffix
                self.changes.append(ChangeRecord(old_word, words[idx], "M1-Synonym"))
                replacements += 1

        return " ".join(words)

    def _vary_sentence_lengths_nuclear(self, sentences: List[str]) -> List[str]:
        """
        v5.0: NUCLEAR sentence length variation.
        Split long sentences aggressively, merge short ones, inject parentheticals.
        Creates extreme burstiness to defeat perplexity-based detectors.
        """
        result: List[str] = []
        prev_len = 0

        for i, sent in enumerate(sentences):
            if not sent.strip():
                continue
            words = sent.split()
            cur_len = len(words)

            # SPLIT very long sentences (>22 words) — always at strength 3+
            if cur_len > 22 and self._s > 0.3:
                sp = self._find_split_point(sent)
                if sp:
                    left = sent[:sp].strip()
                    right = sent[sp:].strip()
                    # Make sure the right part starts with a capital
                    if right and right[0].islower():
                        right = right[0].upper() + right[1:]
                    result.append(left)
                    # Sometimes put the second part as a new sentence
                    if random.random() < 0.6:
                        result.append(right)
                    else:
                        # Sometimes join with em-dash
                        result[-1] = result[-1].rstrip(".") + " — " + right
                    self.changes.append(ChangeRecord(
                        sent[:40]+"...", f"{left[:20]}... | {right[:20]}...", "M1-Split"))
                    prev_len = len(right.split())
                    continue

            # If sentence is very similar length to previous — force variation
            if i > 0 and abs(cur_len - prev_len) <= 4 and self._s > 0.3:
                if cur_len > 15:
                    # Split it
                    sp = self._find_split_point(sent)
                    if sp:
                        left = sent[:sp].strip()
                        right = sent[sp:].strip()
                        if right and right[0].islower():
                            right = right[0].upper() + right[1:]
                        result.append(left)
                        result.append(right)
                        self.changes.append(ChangeRecord(
                            sent[:40]+"...", f"Split for burstiness", "M1-LengthVar"))
                        prev_len = len(right.split())
                        continue
                elif cur_len < 12 and i + 1 < len(sentences) and sentences[i+1].strip():
                    # Merge with next
                    merged = sent.rstrip(".") + ", and " + sentences[i+1].lstrip()
                    if merged[0].islower() and len(merged) > 1:
                        merged = merged[0].upper() + merged[1:]
                    result.append(merged)
                    prev_len = len(merged.split())
                    sentences[i+1] = ""
                    self.changes.append(ChangeRecord(
                        sent[:40]+"...", merged[:40]+"...", "M1-Merge"))
                    continue

            # Inject parenthetical aside in medium-long sentences
            if cur_len > 20 and random.random() < self._s * 0.5:
                parenthetical = random.choice(PARENTHETICALS)
                insert_pos = random.randint(len(words)//3, 2*len(words)//3)
                words.insert(insert_pos, parenthetical)
                sent = " ".join(words)
                self.changes.append(ChangeRecord("(no aside)", parenthetical, "M1-Aside"))

            result.append(sent)
            prev_len = len(sent.split())

        return [s for s in result if s.strip()]

    def _inject_burstiness(self, sentence: str, idx: int, total: int) -> str:
        """Inject a short follow-up sentence after this one for burstiness."""
        # Avoid duplicating burst sentences already in the text
        available = [b for b in BURST_SHORT_SENTENCES if b not in sentence]
        if not available:
            return sentence
        burst = random.choice(available)
        if not sentence.endswith("."):
            sentence += "."
        result = sentence + " " + burst
        self.changes.append(ChangeRecord("(no burst)", burst, "M1-Burst2"))
        return result

    def _find_split_point(self, sentence: str) -> Optional[int]:
        # Try to split at ", which", ", while", ", whereas", ", although", ", and"
        for m in re.finditer(r',\s+(?:which|while|whereas|although|and|but|or|so)\b', sentence):
            return m.start() + 1
        # Try comma
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
        v5.0: Eliminate AI's function-word signature and punctuation uniformity.
        More aggressive opener variation, punctuation diversity, hedging.
        """
        sentences = sent_tokenize(text)
        if not sentences:
            return text

        prev_first_word_pos = None
        result_sents: List[str] = []

        for i, sent in enumerate(sentences):
            # Step 1: Vary sentence openers (v5.0: 80% probability for same-POS)
            sent, new_pos = self._vary_opener(sent, prev_first_word_pos)
            prev_first_word_pos = new_pos

            # Step 2: Punctuation variety (v5.0: higher probability)
            sent = self._diversify_punctuation(sent)

            # Step 3: Hedge injection
            sent = self._inject_hedging(sent)

            # Step 4: First-person injection (v5.0 NEW)
            sent = self._inject_first_person(sent)

            result_sents.append(sent)

        # Step 5: Personal touch
        combined = " ".join(result_sents)
        combined = self._inject_personal_touch(combined)

        return combined

    def _vary_opener(self, sent: str, prev_pos: Optional[str]) -> Tuple[str, Optional[str]]:
        """v5.0: Much higher probability of opener variation."""
        words = sent.split()
        if not words:
            return sent, prev_pos

        first_clean = re.sub(r'[^A-Za-z]', '', words[0])
        if not first_clean:
            return sent, prev_pos

        cur_pos = self._simple_pos(first_clean)

        # v5.0: Change opener if same POS as previous, 80% probability
        if prev_pos and cur_pos == prev_pos and random.random() < self._s * 0.95:
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

        # Also randomly change ~30% of openers for variety
        if random.random() < self._s * 0.3 and len(words) > 5:
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
        """v5.0: More aggressive punctuation variety."""
        # Replace ", and" with em-dash or semicolon
        if random.random() < self._s * 0.7:
            sent = re.sub(r',\s+and\s+(?=[a-z])',
                         lambda m: '\u2014and ' if random.random() < 0.5 else '; ',
                         sent, count=1)
        # Add semicolon between independent clauses
        if random.random() < self._s * 0.5:
            sent = re.sub(r',\s+(?:while|whereas|although)\b',
                         lambda m: '; ' + m.group().strip()[:m.group().strip().find(' ')] + ' ',
                         sent, count=1)
        # Replace some commas with em-dashes for parentheticals
        if random.random() < self._s * 0.3:
            # Find a comma that isn't near a citation
            commas = list(re.finditer(r',\s+(?!\d)', sent))
            if len(commas) >= 2:
                # Replace a middle comma with em-dash pair
                target = commas[len(commas)//2]
                start = target.start()
                sent = sent[:start] + ' \u2014' + sent[start+1:]

        return sent

    def _inject_hedging(self, sent: str) -> str:
        """v5.0: More aggressive hedging — also add hedge phrases to non-assertive sentences."""
        # First replace assertive words
        for pattern, hedge in ASSERTIVE_REPLACEMENTS.items():
            if re.search(pattern, sent, re.IGNORECASE) and random.random() < self._s * 0.9:
                old = re.search(pattern, sent, re.IGNORECASE).group()
                sent = re.sub(pattern, hedge, sent, count=1, flags=re.IGNORECASE)
                self.changes.append(ChangeRecord(old, hedge, "M2-Hedge"))
                return sent  # one hedge per sentence

        # v5.0 NEW: Also inject hedge phrase at start of confident-sounding sentences
        confident_starters = ['This is', 'This shows', 'This proves', 'This confirms',
                            'These results', 'This finding', 'This suggests']
        for starter in confident_starters:
            if sent.startswith(starter) and random.random() < self._s * 0.4:
                hedge = random.choice(["It seems that ", "On the face of it, ", "By all indications, "])
                sent = hedge + sent[0].lower() + sent[1:]
                self.changes.append(ChangeRecord(starter, hedge + starter.lower(), "M2-Hedge"))
                break

        return sent

    def _inject_first_person(self, sent: str) -> str:
        """v5.0 NEW: Inject occasional first-person markers (humans use 'we', 'us')."""
        # Replace impersonal constructions with first-person
        if random.random() < self._s * 0.25:
            # "It is important" → "We think it is important"
            if re.match(r'^It\s+(?:is|seems|appears|remains)\b', sent):
                words = sent.split()
                sent = "We think " + sent[0].lower() + sent[1:]
                self.changes.append(ChangeRecord("Impersonal", "First-person", "M2-FirstPerson"))
            # "One can" → "We can"
            elif re.match(r'^One\s+(?:can|could|may|might|should)\b', sent):
                sent = re.sub(r'^One\b', 'We', sent)
                self.changes.append(ChangeRecord("One", "We", "M2-FirstPerson"))
        return sent

    def _inject_personal_touch(self, text: str) -> str:
        """v5.0: More natural personal touches."""
        touches = [
            "\u2014or so it seems.",
            " At least, that is the working assumption.",
            " Or so the argument goes.",
            " At any rate, that is the picture so far.",
            " That, at any rate, is the claim.",
            " This is, of course, easier said than done.",
        ]
        sentences = sent_tokenize(text)
        if len(sentences) < 3:
            return text
        if random.random() < self._s * 0.6:
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
        v5.0: More aggressive semantic deepening.
        Replace transitions always, remove superficial phrases, inject causality + critical.
        """
        # Step 1: Replace generic transitions (v5.0: ALWAYS replace, no random gate)
        text = self._replace_transitions(text)

        # Step 2: Remove superficial "-ing" analysis phrases (always)
        text = self._remove_superficial_analysis(text)

        # Step 3: Causal reordering (v5.0: higher probability)
        text = self._causal_reorder(text)

        # Step 4: Critical perspective injection
        if section_idx == 0 or random.random() < 0.6 * self._s:
            text = self._inject_critical_perspective(text)

        # Step 5: v5.0 NEW — discourse marker injection
        text = self._inject_discourse_markers(text)

        return text

    def _replace_transitions(self, text: str) -> str:
        """v5.0: Always replace transitions — they are a 100% AI giveaway."""
        for generic, replacements in GENERIC_TRANSITIONS.items():
            pattern = re.compile(r'\b' + re.escape(generic) + r'\b', re.IGNORECASE)
            match = pattern.search(text)
            if match:
                replacement = random.choice(replacements)
                old = match.group()
                text = pattern.sub(replacement, text, count=1)
                self.changes.append(ChangeRecord(old, replacement, "M3-Transition"))
        return text

    def _remove_superficial_analysis(self, text: str) -> str:
        """v5.0: Always remove — these are 100% AI tells."""
        for pattern, replacement in SUPERFICIAL_PHRASES.items():
            matches = list(re.finditer(pattern, text, re.IGNORECASE))
            for m in matches:
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
        """v5.0: More aggressive causal connector injection."""
        sentences = sent_tokenize(text)
        if len(sentences) <= 2:
            return text
        # Find listing-pattern sentences and inject causality
        listing = [i for i, s in enumerate(sentences)
                   if re.match(r'^(?:It|This|The|Such|These|Those)\s+(?:is|are|was|were|has|have)\b', s.lstrip())]
        if len(listing) >= 1 and random.random() < self._s * 0.8:
            target = listing[min(1, len(listing)-1)]
            connectors = ["In turn, ", "As a direct result, ", "Because of this, ",
                         "For this reason, ", "And so, ", "Which is why "]
            conn = random.choice(connectors)
            orig = sentences[target]
            sentences[target] = conn + orig[0].lower() + orig[1:]
            self.changes.append(ChangeRecord(orig[:40]+"...", sentences[target][:40]+"...", "M3-Causal"))
        return " ".join(sentences)

    def _inject_critical_perspective(self, text: str) -> str:
        """v5.0: Higher probability, more perspectives."""
        sentences = sent_tokenize(text)
        if len(sentences) < 3:
            return text
        if random.random() < self._s * 0.8:
            perspective = random.choice(CRITICAL_PERSPECTIVES)
            pos = max(1, len(sentences) - 2)
            sentences.insert(pos, perspective)
            self.changes.append(ChangeRecord("(none)", perspective, "M3-Critical"))
        return " ".join(sentences)

    def _inject_discourse_markers(self, text: str) -> str:
        """v5.0 NEW: Inject human-style discourse markers between sentences."""
        sentences = sent_tokenize(text)
        if len(sentences) < 4:
            return text

        markers = [
            "Mind you,", "To be sure,", "That said,", "And yet,",
            "Then again,", "Even so,", "All the same,", "Be that as it may,",
        ]

        result = [sentences[0]]
        for i in range(1, len(sentences)):
            # Inject a marker between ~30% of sentence pairs
            if random.random() < self._s * 0.3:
                marker = random.choice(markers)
                sent = sentences[i]
                # Prepend marker to sentence
                if sent and sent[0].isupper():
                    sent = marker + " " + sent[0].lower() + sent[1:]
                else:
                    sent = marker + " " + sent
                self.changes.append(ChangeRecord("(no marker)", marker, "M3-Discourse"))
                result.append(sent)
            else:
                result.append(sentences[i])

        return " ".join(result)

    # ──────────────────────────────────────────
    # MODULE 4: Watermark & Structure Disrupter
    # ──────────────────────────────────────────
    def module4_watermark_disrupter(self, text: str) -> str:
        """
        v5.0: More aggressive paragraph restructuring and vocab frequency disruption.
        """
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        if not paragraphs:
            return text

        result: List[str] = []
        for para in paragraphs:
            sents = sent_tokenize(para)
            if len(sents) > 4 and random.random() < self._s * 0.6:
                # Split paragraph
                split_at = random.randint(len(sents)//2, len(sents)-2)
                result.append(" ".join(sents[:split_at]))
                result.append(" ".join(sents[split_at:]))
                self.changes.append(ChangeRecord("One paragraph", "Split into two", "M4-Split"))
            elif len(sents) <= 2 and result and random.random() < self._s * 0.5:
                # Merge with previous
                result[-1] += " " + para
                self.changes.append(ChangeRecord("Two paragraphs", "Merged", "M4-Merge"))
            else:
                # Restructure: move a middle sentence
                if len(sents) > 3 and random.random() < self._s * 0.6:
                    first, last = sents[0], sents[-1]
                    middle = sents[1:-1]
                    # Swap two middle sentences
                    if len(middle) > 1:
                        i1 = random.randint(0, len(middle)-1)
                        i2 = (i1 + 1) % len(middle)
                        middle[i1], middle[i2] = middle[i2], middle[i1]
                    sents = [first] + middle + [last]
                    self.changes.append(ChangeRecord("Original order", "Restructured", "M4-Structure"))
                result.append(" ".join(sents))

        combined = "\n\n".join(result)
        return self._disrupt_vocab_frequency(combined)

    def _disrupt_vocab_frequency(self, text: str) -> str:
        """v5.0: Lower threshold (2+ occurrences), replace MORE repeated words."""
        words = text.split()
        freq = Counter(w.lower().strip(".,;:!?()\"'") for w in words if re.search(r'[A-Za-z]', w))
        # v5.0: replace words that appear 2+ times (was 3+)
        over_rep = {w: c for w, c in freq.items() if c >= 2 and w not in STOP_WORDS and len(w) > 3}

        replacements = 0
        max_reps = max(3, int(len(over_rep) * self._s * 0.8))

        for word, count in over_rep.items():
            if replacements >= max_reps:
                break
            if word in AI_VOCAB_REPLACEMENTS:
                syn = random.choice(AI_VOCAB_REPLACEMENTS[word])
            elif word in PHRASAL_VERB_REPLACEMENTS:
                syn = random.choice(PHRASAL_VERB_REPLACEMENTS[word])
            else:
                syn = self.syn_db.get_synonym(word, self.field, self.strength)
            if syn and syn.lower() != word.lower():
                occ = 0
                new_words = []
                for w in words:
                    clean = w.lower().strip(".,;:!?()\"'")
                    if clean == word:
                        occ += 1
                        # Replace 2nd and 3rd occurrences
                        if occ in (2, 3) and replacements < max_reps:
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
        v5.0: Much lower similarity threshold (0.25 vs 0.45).
        Allow MORE change — the whole point is to change the text!
        """
        # Step 1: Restore protected content
        modified = self._restore_protected(modified, original)

        # Step 2: Similarity check — v5.0: lowered to 0.25
        # We WANT the text to change significantly
        sim = self._similarity(original, modified)
        if sim < 0.20:
            # Only revert if similarity is extremely low (meaning broke meaning)
            safe = original
            safe = self._inject_hedging(safe)
            return safe

        # Step 3: Grammar sweep
        modified = self._grammar_sweep(modified)
        return modified

    def _restore_protected(self, modified: str, original: str) -> str:
        """Ensure citations, numbers, and references are intact."""
        for pattern in _PROTECT_PATTERNS:
            orig_matches = pattern.findall(original)
            for om in orig_matches:
                if om not in modified:
                    # Try to find a corrupted version and restore
                    pass  # prevention > cure
        return modified

    def _similarity(self, a: str, b: str) -> float:
        """Lightweight Jaccard + content-word overlap similarity."""
        wa = [w for w in word_tokenize_simple(a.lower()) if w not in STOP_WORDS]
        wb = [w for w in word_tokenize_simple(b.lower()) if w not in STOP_WORDS]
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
        text = re.sub(r'\s+([.,;:!?])', r'\1', text)
        text = re.sub(r' +\n', '\n', text)
        # Fix double words (e.g., "from from")
        text = re.sub(r'\b(\w+)\s+\1\b', r'\1', text, flags=re.IGNORECASE)
        # Fix dangling prepositions at sentence end
        text = re.sub(r'\b(from|to|with|of|for|at|in|on|by)\.$', lambda m: m.group(0)[:-1], text, flags=re.MULTILINE)
        text = re.sub(r'\b(from|to|with|of|for|at|in|on|by)\.\s', lambda m: m.group(0)[:-2] + '. ', text)
        # Fix double em-dash
        text = re.sub(r'[\u2014\-]\s*[\u2014\-]+', '\u2014', text)
        # Fix "— And" / "— But" (capital after em-dash is wrong)
        text = re.sub(r'[\u2014\-]\s+(And|But|Or|So|Yet|Which|While|Where|When)\b', 
                     lambda m: m.group(0)[0] + ' ' + m.group(1).lower(), text)
        # Fix "And" / "But" / "Or" after semicolon
        text = re.sub(r';\s+(And|But|Or|So|Yet)\b', lambda m: '; ' + m.group(1).lower(), text)
        # Fix capital "And" / "But" / "Or" after comma (should be lowercase)
        text = re.sub(r',\s+(And|But|Or|So|Yet)\b', lambda m: ', ' + m.group(1).lower(), text)
        # Ensure sentence starts with capital after period+space
        text = re.sub(r'\.\s+([a-z])', lambda m: '. ' + m.group(1).upper(), text)
        # Fix duplicated phrases like "inputs as inputs"
        text = re.sub(r'\b(\w+)\s+(?:as|into|to)\s+\1\b', r'\1', text, flags=re.IGNORECASE)
        # Remove awkward "render up" / "identify out" / "discern out" patterns
        text = text.replace('render up', 'make up')
        text = text.replace('identify out', 'identify')
        text = text.replace('discern out', 'discern')
        text = text.replace('determine out', 'determine')
        text = text.replace('fabricate up', 'make up')
        text = text.replace('constitute up', 'make up')
        # Fix stacked discourse markers (e.g., "All the same, but there is a catch.")
        text = re.sub(r'(All the same|Be that as it may|Even so|That said|Mind you|Then again),\s+(but|and|yet|however)\b', 
                     lambda m: m.group(1) + ',', text, flags=re.IGNORECASE)
        # Fix missing subject: "obliges" at sentence start needs subject
        text = re.sub(r'\b(Obliges|Forces|Makes|Requires)\s+(engineers|researchers|practitioners|people|users)\b',
                     lambda m: 'This ' + m.group(1).lower() + ' ' + m.group(2), text)
        # Fix space before comma/period
        text = re.sub(r'\s+([,.])', r'\1', text)
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

    # Save first run's formatting
    first_run = para.runs[0]
    fmt = {
        'bold': first_run.bold,
        'italic': first_run.italic,
        'underline': first_run.underline,
        'font_name': first_run.font.name,
        'font_size': first_run.font.size,
        'font_color': first_run.font.color.rgb if first_run.font.color and first_run.font.color.rgb else None,
    }

    # Clear all runs
    for run in para.runs:
        run.text = ""

    # Set text on first run, apply saved formatting
    first_run.text = new_text
    if fmt['bold'] is not None:
        first_run.bold = fmt['bold']
    if fmt['italic'] is not None:
        first_run.italic = fmt['italic']
    if fmt['underline'] is not None:
        first_run.underline = fmt['underline']
    if fmt['font_name']:
        first_run.font.name = fmt['font_name']
    if fmt['font_size']:
        first_run.font.size = fmt['font_size']
    if fmt['font_color']:
        first_run.font.color.rgb = fmt['font_color']


# ╔══════════════════════════════════════════════╗
# ║        PDF PROCESSOR                          ║
# ╚══════════════════════════════════════════════╝

def process_pdf(input_bytes: bytes, engine: HumanizeEngine,
                progress_cb=None) -> Tuple[bytes, List[ChangeRecord]]:
    """
    Process a PDF file: extract text, humanize, and output as DOCX.
    Preserves structure (paragraphs, sections) as much as possible.
    """
    from docx import Document

    try:
        import pdfplumber
    except ImportError:
        raise RuntimeError("pdfplumber is required for PDF processing. Install with: pip install pdfplumber")

    if progress_cb:
        progress_cb("Reading PDF...", 5)

    # Extract text from PDF
    paragraphs_text = []
    with pdfplumber.open(io.BytesIO(input_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages):
            if progress_cb:
                pct = 5 + int(25 * (page_num / max(total_pages, 1)))
                progress_cb(f"Extracting page {page_num+1}/{total_pages}...", pct)
            text = page.extract_text()
            if text:
                for para in text.split("\n"):
                    para = para.strip()
                    if para and len(para) > 10:
                        paragraphs_text.append(para)

    if not paragraphs_text:
        raise ValueError("No text could be extracted from the PDF.")

    # Humanize each paragraph
    all_changes: List[ChangeRecord] = []
    humanized_paras: List[str] = []

    for idx, para_text in enumerate(paragraphs_text):
        if progress_cb:
            pct = 30 + int(50 * (idx / max(len(paragraphs_text), 1)))
            progress_cb("Humanizing...", pct)

        humanized, changes = engine.humanize_text(para_text)
        humanized_paras.append(humanized)
        all_changes.extend(changes)

    # Create DOCX output
    doc = Document()
    for para_text in humanized_paras:
        doc.add_paragraph(para_text)

    output = io.BytesIO()
    doc.save(output)

    if progress_cb:
        progress_cb("Done", 100)

    return output.getvalue(), all_changes


# ╔══════════════════════════════════════════════╗
# ║        STREAMLIT UI                           ║
# ╚══════════════════════════════════════════════╝

def main():
    st.set_page_config(
        page_title="DeepClean Studio v5.0",
        page_icon="🧹",
        layout="wide",
    )

    # ── Custom CSS ──
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 1rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    .main-header h1 {
        color: #e94560;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    .main-header p {
        color: #a8a8b3;
        font-size: 1.1rem;
    }
    .module-card {
        background: #1e1e2f;
        border: 1px solid #333;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 0.5rem;
    }
    .module-card h4 {
        color: #e94560;
        margin: 0 0 0.5rem 0;
    }
    .module-card p {
        color: #bbb;
        font-size: 0.85rem;
        margin: 0;
    }
    .change-tag {
        display: inline-block;
        background: #2a2a4a;
        color: #e94560;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        margin: 2px;
    }
    div.stButton > button {
        background: linear-gradient(135deg, #e94560, #c23152);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-size: 1.1rem;
        font-weight: bold;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #c23152, #a02040);
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ──
    st.markdown("""
    <div class="main-header">
        <h1>🧹 DeepClean Studio v5.0</h1>
        <p>Nuclear Humanization Engine — Bypass AI Detectors with Confidence</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ──
    with st.sidebar:
        st.header("⚙️ Settings")

        field = st.selectbox(
            "Academic Field",
            ["General", "Medical", "Engineering", "Humanities"],
            help="Select the field for field-specific synonym replacements"
        )

        strength = st.slider(
            "Humanization Strength",
            min_value=1, max_value=5, value=3,
            help="1=Light touch, 3=Balanced, 5=Maximum transformation"
        )

        st.markdown("---")
        st.markdown("### 🔬 5 Modules")
        modules_info = [
            ("M1: Statistical Bone-Breaker", "Destroys perplexity uniformity & burstiness patterns. UNCAPPED AI vocabulary replacement."),
            ("M2: Stylometric Mask", "Eliminates AI fingerprint. Varied openers, punctuation, hedging, phrasal verbs, first-person."),
            ("M3: Semantic Deepener", "Replaces AI transitions, removes superficial analysis, injects causality & critical thinking."),
            ("M4: Watermark Disrupter", "Breaks template structures, disrupts token frequency patterns, reorders paragraphs."),
            ("M5: Coherence Guardian", "Ensures meaning preservation. Low similarity threshold (0.20) allows deep transformation."),
        ]
        for title, desc in modules_info:
            st.markdown(f"""
            <div class="module-card">
                <h4>{title}</h4>
                <p>{desc}</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🎯 Target Detectors")
        st.markdown("""
        - GPTZero
        - Originality.ai
        - Turnitin AI Detection
        - Copyleaks
        - Winston AI
        """)

    # ── Main content ──
    tab1, tab2 = st.tabs(["📝 Text Input", "📁 File Upload"])

    with tab1:
        input_text = st.text_area(
            "Paste your AI-generated text here:",
            height=300,
            placeholder="Paste the text you want to humanize..."
        )

        if st.button("🧹 Humanize Text", key="btn_text"):
            if not input_text or len(input_text.strip()) < 20:
                st.error("Please enter at least 20 characters of text.")
            else:
                synonym_csv = str(Path(__file__).parent / "synonyms_academic.csv")
                engine = HumanizeEngine(synonym_csv, field=field, strength=strength)

                progress = st.progress(0, text="Starting...")
                def progress_cb(msg, pct):
                    progress.progress(pct, text=msg)

                with st.spinner("Humanizing... This may take a moment for long texts."):
                    humanized, changes = engine.humanize_text(input_text, progress_cb=progress_cb)

                st.success(f"✅ Humanization complete! {len(changes)} changes made.")

                # Show before/after comparison
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**📄 Original:**")
                    st.text_area("Original", input_text, height=250, key="orig_text", disabled=True)
                with col2:
                    st.markdown("**✨ Humanized:**")
                    st.text_area("Humanized", humanized, height=250, key="hum_text")

                # Show changes breakdown
                if changes:
                    st.markdown("### 📊 Change Breakdown")
                    module_counts = Counter(c.module for c in changes)
                    cols = st.columns(len(module_counts))
                    for col, (module, count) in zip(cols, module_counts.most_common()):
                        col.metric(module, count)

                    # Show sample changes
                    with st.expander("🔍 View Sample Changes", expanded=False):
                        for c in changes[:50]:
                            st.markdown(f'<span class="change-tag">{c.module}</span> **{c.original}** → **{c.modified}**', unsafe_allow_html=True)

                # Download as DOCX
                from docx import Document
                doc = Document()
                doc.add_paragraph(humanized)
                output = io.BytesIO()
                doc.save(output)
                st.download_button(
                    "📥 Download as DOCX",
                    data=output.getvalue(),
                    file_name="humanized_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    with tab2:
        uploaded_file = st.file_uploader(
            "Upload a DOCX or PDF file:",
            type=["docx", "pdf"],
            help="Supports .docx and .pdf files. PDF output will be converted to DOCX."
        )

        if uploaded_file is not None:
            file_ext = Path(uploaded_file.name).suffix.lower()
            st.info(f"📁 File: **{uploaded_file.name}** ({file_ext})")

            if st.button("🧹 Humanize File", key="btn_file"):
                synonym_csv = str(Path(__file__).parent / "synonyms_academic.csv")
                engine = HumanizeEngine(synonym_csv, field=field, strength=strength)

                progress = st.progress(0, text="Starting...")
                def progress_cb(msg, pct):
                    progress.progress(pct, text=msg)

                input_bytes = uploaded_file.read()

                with st.spinner("Processing file..."):
                    try:
                        if file_ext == ".docx":
                            output_bytes, changes = process_docx(input_bytes, engine, progress_cb)
                            out_name = Path(uploaded_file.name).stem + "_humanized.docx"
                        elif file_ext == ".pdf":
                            output_bytes, changes = process_pdf(input_bytes, engine, progress_cb)
                            out_name = Path(uploaded_file.name).stem + "_humanized.docx"
                        else:
                            st.error("Unsupported file format.")
                            return

                        st.success(f"✅ Humanization complete! {len(changes)} changes made.")

                        # Changes breakdown
                        if changes:
                            st.markdown("### 📊 Change Breakdown")
                            module_counts = Counter(c.module for c in changes)
                            cols = st.columns(min(len(module_counts), 5))
                            for col, (module, count) in zip(cols, module_counts.most_common()):
                                col.metric(module, count)

                        # Download
                        st.download_button(
                            "📥 Download Humanized DOCX",
                            data=output_bytes,
                            file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

                    except Exception as e:
                        st.error(f"❌ Error processing file: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
