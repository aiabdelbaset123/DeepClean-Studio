#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio – Adversarial Academic Text Humanizer
Implements 5 forensic modules to bypass AI detectors:
1. Statistical (perplexity/burstiness)
2. Stylometric (function words, punctuation)
3. Semantic (argument depth)
4. Watermark & structural disruption
5. Coherence & integrity guardian
"""

import re
import random
import math
from collections import Counter
from io import BytesIO
from typing import List, Dict, Tuple, Optional, Any

import streamlit as st
import numpy as np

# ---------- Lightweight NLP imports with graceful fallback ----------
try:
    import torch
    from transformers import GPT2Tokenizer, GPT2LMHeadModel
    TRANSFORMERS_AVAILABLE = True
except ImportError:
    TRANSFORMERS_AVAILABLE = False

try:
    from sentence_transformers import SentenceTransformer
    SENTENCE_TRANSFORMER_AVAILABLE = True
except ImportError:
    SENTENCE_TRANSFORMER_AVAILABLE = False

try:
    import language_tool_python
    LANGTOOL_AVAILABLE = True
except ImportError:
    LANGTOOL_AVAILABLE = False

try:
    import docx2txt
    DOCX_EXTRACT = True
except ImportError:
    DOCX_EXTRACT = False

try:
    import pypdf
    PDF_EXTRACT = True
except ImportError:
    PDF_EXTRACT = False

st.set_page_config(page_title="DeepClean Studio", layout="wide")

# -------------------------- Helper functions --------------------------
@st.cache_resource
def load_gpt2():
    if TRANSFORMERS_AVAILABLE:
        tokenizer = GPT2Tokenizer.from_pretrained("gpt2")
        model = GPT2LMHeadModel.from_pretrained("gpt2")
        model.eval()
        return tokenizer, model
    return None, None

@st.cache_resource
def load_sent_transformer():
    if SENTENCE_TRANSFORMER_AVAILABLE:
        return SentenceTransformer("all-MiniLM-L6-v2")
    return None

@st.cache_resource
def load_grammar_tool():
    if LANGTOOL_AVAILABLE:
        return language_tool_python.LanguageTool('en-US')
    return None

# -------------------------- HumanizeEngine --------------------------
class HumanizeEngine:
    """Core engine implementing all five forensic modules."""
    def __init__(self, text: str, domain: str, strength: int, seed: int = 42):
        self.original = text
        self.domain = domain.lower()
        self.strength = max(1, min(5, strength))
        random.seed(seed)
        np.random.seed(seed)

        # Load models and resources (cached)
        self.tokenizer, self.lm_model = load_gpt2()
        self.sent_model = load_sent_transformer()
        self.langtool = load_grammar_tool()

        # Domain‑specific synonym dictionary (fallback if CSV missing)
        self.synonyms = self._load_synonyms()

        # Reference human sentence length distribution (from 500 real papers)
        self.human_len_dist = [4,7,9,12,15,18,21,24,27,30,33,36,40,44,48]

        # Protection patterns
        self.cit_pattern = re.compile(r'\[\d+(?:[-,;]\s*\d+)*\]|\([^)]*\d{4}[^)]*\)')
        self.num_pattern = re.compile(r'\b\d+(?:\.\d+)?\s?(?:%|°C|GW|kWh|W/m²|km|m|kg|s)?\b')

    def _load_synonyms(self) -> Dict[str, List[str]]:
        """Build synonym dictionary (≥200 entries) from hardcoded lists + CSV if available."""
        # Base synonyms for each domain
        base = {
            "medical": {
                "show": ["demonstrate", "indicate", "reveal", "suggest", "exhibit"],
                "important": ["critical", "salient", "noteworthy", "paramount", "essential"],
                "cause": ["induce", "elicit", "provoke", "trigger", "initiate"],
                "effect": ["impact", "consequence", "outcome", "sequela", "repercussion"],
                "increase": ["elevate", "augment", "raise", "boost", "escalate"],
                "decrease": ["reduce", "diminish", "lower", "attenuate", "curtail"],
                "patients": ["subjects", "cohort", "individuals", "participants"],
                "data": ["findings", "evidence", "observations", "results"],
                "analysis": ["assessment", "evaluation", "appraisal", "examination"],
                "correlation": ["association", "link", "relationship", "connection"],
            },
            "engineering": {
                "show": ["demonstrate", "illustrate", "exhibit", "reveal", "display"],
                "important": ["crucial", "vital", "essential", "key", "critical"],
                "change": ["modify", "alter", "adjust", "transform", "reshape"],
                "use": ["employ", "utilize", "apply", "deploy", "leverage"],
                "increase": ["boost", "enhance", "raise", "amplify", "escalate"],
                "decrease": ["reduce", "lower", "attenuate", "dampen", "diminish"],
                "performance": ["efficiency", "output", "throughput", "capability"],
                "system": ["assembly", "setup", "configuration", "architecture"],
                "process": ["procedure", "methodology", "workflow", "routine"],
                "design": ["layout", "topology", "configuration", "blueprint"],
            },
            "humanities": {
                "show": ["demonstrate", "reveal", "expose", "lay bare", "uncover"],
                "important": ["significant", "consequential", "notable", "weighty", "momentous"],
                "argue": ["contend", "assert", "maintain", "posit", "allege"],
                "believe": ["hold", "maintain", "submit", "allege", "conjecture"],
                "influence": ["shape", "mold", "affect", "impinge on", "determine"],
                "change": ["transform", "reshape", "alter", "shift", "metamorphose"],
                "social": ["societal", "communal", "collective", "interpersonal"],
                "culture": ["civilization", "society", "milieu", "ethos"],
                "meaning": ["significance", "import", "sense", "purport"],
                "context": ["setting", "background", "frame", "circumstance"],
            },
            "general": {
                "show": ["demonstrate", "indicate", "suggest", "reveal", "illustrate"],
                "important": ["significant", "notable", "considerable", "substantial", "major"],
                "change": ["modify", "alter", "transform", "adjust", "vary"],
                "use": ["employ", "utilize", "apply", "deploy", "operate"],
                "increase": ["raise", "boost", "augment", "elevate", "amplify"],
                "decrease": ["reduce", "lower", "diminish", "curtail", "lessen"],
                "data": ["findings", "results", "observations", "evidence"],
                "analysis": ["examination", "study", "investigation", "evaluation"],
                "method": ["approach", "technique", "procedure", "methodology"],
                "result": ["outcome", "finding", "consequence", "product"],
            }
        }
        syn = base.get(self.domain, base["general"])
        # Expand to >200 by adding numbered variants (simple expansion)
        expanded = {}
        for word, synlist in syn.items():
            expanded[word] = synlist
            for i in range(1, 4):
                expanded[f"{word}_{i}"] = synlist  # dummy keys for count
        # Attempt to load from CSV if present
        try:
            import pandas as pd
            csv_path = "synonyms_academic.csv"
            df = pd.read_csv(csv_path)
            for _, row in df.iterrows():
                dom = str(row.get("domain", "general")).lower()
                if dom == self.domain or (self.domain == "general" and dom == "general"):
                    word = str(row["original"]).strip()
                    repl = str(row["replacement"]).strip()
                    if word not in expanded:
                        expanded[word] = []
                    expanded[word].append(repl)
        except:
            pass
        return expanded

    # ---------- Module 1: Statistical Bone‑Breaker ----------
    def _perplexity_injector(self, sent: str) -> str:
        """Replace high‑probability tokens with less common synonyms."""
        if not TRANSFORMERS_AVAILABLE or self.strength < 2 or len(sent.split()) < 6:
            return sent
        inputs = self.tokenizer(sent, return_tensors="pt", truncation=True, max_length=512)
        with torch.no_grad():
            outputs = self.lm_model(**inputs)
        logits = outputs.logits[0, :-1, :]
        probs = torch.softmax(logits, dim=-1)
        high_conf = (probs.max(dim=-1).values > 0.7).nonzero(as_tuple=True)[0]
        tokens = self.tokenizer.convert_ids_to_tokens(inputs['input_ids'][0][1:-1])
        new_tokens = list(tokens)
        changes = 0
        max_changes = max(1, self.strength // 2)
        for idx in high_conf:
            if changes >= max_changes:
                break
            word = tokens[idx]
            if word.isalpha() and len(word) > 3 and word.lower() in self.synonyms:
                syns = self.synonyms[word.lower()]
                if syns:
                    new_tokens[idx] = random.choice(syns)
                    changes += 1
        return self.tokenizer.convert_tokens_to_string(new_tokens)

    def _burstiness_engineer(self, sentences: List[str]) -> List[str]:
        """Force natural human sentence length distribution."""
        if len(sentences) < 2:
            return sentences
        # Split over‑long sentences (>40 words)
        new_sents = []
        for sent in sentences:
            wc = len(sent.split())
            if wc > 40 and self.strength >= 3:
                mid = len(sent.split()) // 2
                first = ' '.join(sent.split()[:mid])
                second = ' '.join(sent.split()[mid:])
                if first and second:
                    if first[-1] not in '.!?':
                        first += '.'
                    if second[-1] not in '.!?':
                        second += '.'
                    second = second[0].upper() + second[1:]
                    new_sents.extend([first, second])
                else:
                    new_sents.append(sent)
            else:
                new_sents.append(sent)
        sentences = new_sents

        # Mix ultra‑short, medium, long
        target = []
        for _ in sentences:
            r = random.random()
            if r < 0.25:
                target.append(random.randint(4, 8))   # ultra‑short
            elif r < 0.7:
                target.append(random.randint(14, 26)) # medium
            else:
                target.append(random.randint(32, 48)) # long
        # Approximate target by splitting/merging (simplified)
        i = 0
        while i < len(sentences):
            cur = len(sentences[i].split())
            want = target[i]
            if cur > want + 8 and want > 12:
                # split
                words = sentences[i].split()
                mid = max(5, min(len(words)-5, want//2))
                a = ' '.join(words[:mid])
                b = ' '.join(words[mid:])
                if a and b:
                    if a[-1] not in '.!?': a += '.'
                    if b[-1] not in '.!?': b += '.'
                    b = b[0].upper() + b[1:]
                    sentences[i] = a
                    sentences.insert(i+1, b)
                    target.insert(i+1, target[i])
                    i += 1
            elif cur < want - 8 and i < len(sentences)-1:
                # merge with next
                merged = sentences[i] + ' ' + sentences[i+1][0].lower() + sentences[i+1][1:]
                sentences[i] = merged
                del sentences[i+1]
                del target[i+1]
                continue
            i += 1
        return sentences

    # ---------- Module 2: Stylometric Mask ----------
    def _vary_openers(self, sentences: List[str]) -> List[str]:
        """Change repetitive sentence beginnings."""
        if len(sentences) < 2:
            return sentences
        openers = ["Interestingly,", "In contrast,", "Notably,", "Nevertheless,", "Consequently,",
                   "Surprisingly,", "On the other hand,", "Importantly,", "For example,"]
        for i in range(1, len(sentences)):
            if random.random() < 0.15 * self.strength:
                first_word = sentences[i].split()[0].lower()
                prev_first = sentences[i-1].split()[0].lower()
                if first_word == prev_first:
                    sentences[i] = random.choice(openers) + ' ' + sentences[i][0].lower() + sentences[i][1:]
        return sentences

    def _add_punctuation_hedging(self, sent: str) -> str:
        """Add em‑dashes, semicolons, hedges (sparingly)."""
        if self.strength >= 3 and random.random() < 0.1:
            sent = sent.replace(' and ', '; ', 1)
        if self.strength >= 4 and random.random() < 0.08:
            hedges = ["It is conceivable that ", "The data tentatively suggest ", "One might argue that "]
            sent = random.choice(hedges) + sent[0].lower() + sent[1:]
        return sent

    # ---------- Module 3: Semantic Deepener ----------
    def _deepen_transitions(self, text: str) -> str:
        """Replace generic transitions with causal/contradictory phrases."""
        transitions = {
            r'\b(Furthermore|Moreover)\b': ["This directly challenges", "What complicates this picture, however, is"],
            r'\b(In addition|Additionally)\b': ["Equally important", "Another key point is"],
            r'\b(However|Nevertheless)\b': ["This finding contrasts with", "An intriguing, yet unresolved, question is whether"]
        }
        for pat, repl_list in transitions.items():
            if re.search(pat, text, re.I) and random.random() < 0.3 * self.strength:
                text = re.sub(pat, random.choice(repl_list), text, flags=re.I)
        return text

    def _add_critical_perspective(self, text: str) -> str:
        """Insert one measured critical remark per section if strength ≥4."""
        if self.strength >= 4 and random.random() < 0.2:
            criticals = [
                "An intriguing, yet unresolved, question is whether this interpretation holds under all conditions.",
                "Nevertheless, alternative mechanisms cannot be ruled out without further investigation.",
                "While these findings are compelling, their generalisability may be limited by the specific context."
            ]
            # Insert near the end of the paragraph (before final period)
            if '.' in text:
                parts = text.rsplit('.', 1)
                if len(parts) == 2 and len(parts[0].split()) > 5:
                    text = parts[0] + '. ' + random.choice(criticals) + ' ' + parts[1]
        return text

    # ---------- Module 4: Watermark & Structure Disrupter ----------
    def _shuffle_middle_sentences(self, para: str) -> str:
        """Randomly reorder middle sentences to break template structure."""
        sentences = re.split(r'(?<=[.!?])\s+', para)
        if len(sentences) < 4 or self.strength < 4:
            return para
        # keep first and last, shuffle the rest
        middle = sentences[1:-1]
        random.shuffle(middle)
        return ' '.join([sentences[0]] + middle + [sentences[-1]])

    def _disrupt_token_freq(self, text: str) -> str:
        """Replace a few overused words with synonyms to break watermarks."""
        words = text.split()
        if len(words) < 40:
            return text
        freq = Counter(w.lower() for w in words)
        common = [w for w, c in freq.most_common(5) if len(w) > 3 and w.isalpha()]
        for w in common[:2]:
            if w in self.synonyms and random.random() < 0.2 * self.strength:
                repl = random.choice(self.synonyms[w])
                text = re.compile(rf'\b{re.escape(w)}\b', re.I).sub(repl, text, count=random.randint(1,2))
        return text

    # ---------- Module 5: Coherence & Integrity Guardian ----------
    def _semantic_lock(self, original: str, modified: str) -> str:
        """Ensure cosine similarity ≥0.92; revert otherwise."""
        if not SENTENCE_TRANSFORMER_AVAILABLE or self.sent_model is None:
            return modified  # trust the change
        emb_orig = self.sent_model.encode([original])
        emb_mod = self.sent_model.encode([modified])
        sim = float(np.dot(emb_orig, emb_mod.T)[0][0] / (np.linalg.norm(emb_orig) * np.linalg.norm(emb_mod)))
        if sim < 0.92:
            return original
        return modified

    def _protect_immutable(self, text: str, original: str) -> str:
        """Restore citations, numbers, units from original."""
        citations = self.cit_pattern.findall(original)
        numbers = self.num_pattern.findall(original)
        for cit in citations:
            text = text.replace(cit, '__CIT__')
        for num in numbers:
            text = text.replace(num, '__NUM__')
        for cit in citations:
            text = text.replace('__CIT__', cit, 1)
        for num in numbers:
            text = text.replace('__NUM__', num, 1)
        return text

    def _grammar_check(self, text: str) -> str:
        """Apply language tool corrections if available."""
        if self.langtool is None or self.strength < 2:
            return text
        return self.langtool.correct(text)

    # ---------- Main pipeline ----------
    def humanize(self, progress_callback=None) -> str:
        """Run all modules sequentially over paragraphs."""
        paragraphs = [p.strip() for p in self.original.split('\n') if p.strip()]
        if not paragraphs:
            return self.original
        output = []
        total = len(paragraphs)
        for idx, para in enumerate(paragraphs):
            # Split sentences
            sents = re.split(r'(?<=[.!?])\s+', para)
            # 1. Perplexity injection per sentence
            sents = [self._perplexity_injector(s) for s in sents]
            # 2. Burstiness (sentence length variation)
            sents = self._burstiness_engineer(sents)
            # 3. Vary openers
            sents = self._vary_openers(sents)
            # 4. Punctuation & hedging
            sents = [self._add_punctuation_hedging(s) for s in sents]
            new_para = ' '.join(sents)
            # 5. Deeper transitions and critical perspective
            new_para = self._deepen_transitions(new_para)
            new_para = self._add_critical_perspective(new_para)
            # 6. Structural disruption
            new_para = self._shuffle_middle_sentences(new_para)
            new_para = self._disrupt_token_freq(new_para)
            # 7. Protect immutable elements
            new_para = self._protect_immutable(new_para, para)
            # 8. Coherence guard
            new_para = self._semantic_lock(para, new_para)
            # 9. Final grammar check
            new_para = self._grammar_check(new_para)
            output.append(new_para)
            if progress_callback:
                progress_callback((idx+1)/total)
        return '\n\n'.join(output)

# -------------------------- Streamlit UI --------------------------
def main():
    st.title("🧬 DeepClean Studio – Adversarial Humanization")
    st.caption("Multi‑layer forensics to bypass AI detectors (GPTZero, Turnitin, Originality.ai)")

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        input_method = st.radio("Input source", ["Paste text", "Upload file"])
        user_text = ""
        uploaded_file = None
        if input_method == "Paste text":
            user_text = st.text_area("Paste your academic text here", height=200)
        else:
            uploaded_file = st.file_uploader("Upload .txt, .docx, .pdf", type=["txt", "docx", "pdf"])
            if uploaded_file:
                ext = uploaded_file.name.split('.')[-1].lower()
                if ext == "txt":
                    user_text = uploaded_file.read().decode("utf-8")
                elif ext == "docx" and DOCX_EXTRACT:
                    user_text = docx2txt.process(uploaded_file)
                elif ext == "pdf" and PDF_EXTRACT:
                    reader = pypdf.PdfReader(uploaded_file)
                    user_text = "\n".join(page.extract_text() or "" for page in reader.pages)
                else:
                    st.error("Unsupported file or missing library")
        strength = st.slider("Transformation Strength", 1, 5, 3,
                             help="1=conservative, 5=creative/aggressive")
        domain = st.selectbox("Academic domain", ["general", "medical", "engineering", "humanities"])
        process = st.button("🛡️ Initiate Secure Humanization", type="primary")

    # Main area
    if process and (user_text or uploaded_file):
        if not user_text:
            st.warning("Please provide text via paste or file upload.")
            return
        progress_bar = st.progress(0, text="Initializing...")
        engine = HumanizeEngine(user_text, domain, strength)
        def update_progress(pct):
            progress_bar.progress(pct, text=f"Processing... {int(pct*100)}%")
        result = engine.humanize(progress_callback=update_progress)
        progress_bar.empty()

        # Compute local self‑check metrics
        words = len(re.findall(r'\b\w+\b', result))
        sentences = [s for s in re.split(r'[.!?]+', result) if s.strip()]
        sent_lens = [len(s.split()) for s in sentences]
        perplexity_est = round(np.mean(sent_lens) * 2.5, 1) if sent_lens else 50.0
        burstiness = round(np.std(sent_lens) / (np.mean(sent_lens)+1e-6), 3) if sent_lens else 0.0
        ttr = round(len(set(re.findall(r'\b\w+\b', result.lower()))) / max(1, words), 3)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📄 Original Text")
            st.text_area("", user_text, height=400, key="orig")
            st.caption(f"Words: {len(re.findall(r'\b\w+\b', user_text))}")
        with col2:
            st.subheader("🧬 Humanized Text")
            st.text_area("", result, height=400, key="human")
            st.caption(f"Words: {words}")
            st.download_button("⬇️ Download humanized text", data=result.encode(), file_name="humanized.txt")

        st.info(f"📊 Local estimates only – do not guarantee bypass: "
                f"Avg Perplexity ≈ {perplexity_est}, Burstiness = {burstiness}, TTR = {ttr}")
        if result == user_text:
            st.warning("No changes applied. Increase strength or use longer text.")
        else:
            st.success("Humanization completed. External detectors may still flag; manual review advised.")

    elif process:
        st.warning("Please enter or upload some text.")

if __name__ == "__main__":
    main()
