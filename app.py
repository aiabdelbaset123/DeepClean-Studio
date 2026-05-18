#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - النسخة النهائية المُحسَّنة والمُختبرة
تحويل النصوص الأكاديمية من نمط الذكاء الاصطناعي إلى طابع بشري خبير
مع تصدير إلى Word، تدقيق لغوي، وعرض تفصيلي للتغييرات.
"""

import math
import random
import re
import difflib
from collections import Counter
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import docx2txt
import nltk
import numpy as np
import pandas as pd
import pypdf
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from nltk.tokenize import sent_tokenize, word_tokenize
from sentence_transformers import SentenceTransformer, util

# =============================================================================
# إعداد موارد NLTK
# =============================================================================
@st.cache_resource
def download_nltk_resources() -> None:
    resources = {
        "tokenizers/punkt_tab": "punkt_tab",
        "tokenizers/punkt": "punkt",
        "corpora/wordnet": "wordnet",
        "corpora/stopwords": "stopwords",
        "taggers/averaged_perceptron_tagger_eng": "averaged_perceptron_tagger_eng",
    }
    for path, name in resources.items():
        try:
            nltk.data.find(path)
        except LookupError:
            nltk.download(name, quiet=True)

download_nltk_resources()

# =============================================================================
# النموذج الدلالي
# =============================================================================
@st.cache_resource
def load_semantic_model() -> SentenceTransformer:
    return SentenceTransformer("all-MiniLM-L6-v2")

# =============================================================================
# قاموس المرادفات
# =============================================================================
@st.cache_data
def load_synonym_dictionary(filepath: str = "synonyms_academic.csv") -> pd.DataFrame:
    try:
        df = pd.read_csv(filepath, encoding="utf-8")
        if {"domain", "original", "replacement"}.issubset(df.columns):
            return df
        else:
            st.error("ملف المرادفات يجب أن يحتوي على الأعمدة: domain, original, replacement")
            return pd.DataFrame(columns=["domain", "original", "replacement"])
    except FileNotFoundError:
        st.warning("ملف synonyms_academic.csv غير موجود. سيتم استخدام قاموس افتراضي.")
        default_data = {
            "domain": ["medical","medical","engineering","engineering","humanities","humanities","medical","engineering","humanities"],
            "original": ["patient","treatment","system","design","culture","history","diagnosis","algorithm","philosophy"],
            "replacement": [
                "individual receiving medical care","therapeutic intervention",
                "integrated framework","architectural conception",
                "sociocultural paradigm","historical trajectory",
                "clinical assessment","computational procedure",
                "school of thought"
            ]
        }
        return pd.DataFrame(default_data)

# =============================================================================
# الكلاس الرئيسي
# =============================================================================
class DeepCleanEngine:
    def __init__(self, domain: str, intensity: int, text: str) -> None:
        self.domain = domain
        self.intensity = intensity
        self.original_text = text
        self.semantic_model = load_semantic_model()
        self.synonym_df = load_synonym_dictionary()

        self.domain_synonyms: Dict[str, str] = {}
        if not self.synonym_df.empty:
            domain_df = self.synonym_df[self.synonym_df["domain"] == self.domain]
            for _, row in domain_df.iterrows():
                self.domain_synonyms[row["original"].lower()] = row["replacement"].lower()
        if self.domain == "general" and not self.synonym_df.empty:
            for _, row in self.synonym_df.iterrows():
                self.domain_synonyms[row["original"].lower()] = row["replacement"].lower()

        self.reference_lengths = [15,22,18,30,25,19,28,16,21,27,23,17,20,24,26,19,22,31,14,29]

        self.transition_words = {
            "medical": ["furthermore","moreover","conversely","notably","specifically"],
            "engineering": ["additionally","in contrast","consequently","accordingly","particularly"],
            "humanities": ["moreover","on the other hand","thus","indeed","above all"],
            "general": ["furthermore","however","therefore","for instance","in particular"]
        }

        self.causal_links = {
            "medical": ["due to","as a result of","leading to","stemming from"],
            "engineering": ["caused by","resulting in","attributed to","driven by"],
            "humanities": ["owing to","consequently","as a consequence","thereby"],
            "general": ["because of","hence","thus","accordingly"]
        }
        self.contrast_links = {
            "medical": ["whereas","in contrast","conversely","on the contrary"],
            "engineering": ["however","on the other hand","alternatively","in opposition"],
            "humanities": ["nevertheless","nonetheless","yet","contrarily"],
            "general": ["but","however","although","in spite of"]
        }

        self.citation_pattern = re.compile(r"\[[\d,\-; ]+\]|\([^)]*\d{4}[^)]*\)")
        self.number_pattern = re.compile(r"\b\d+(?:\.\d+)?%?\b")

    def _get_synonym(self, word: str) -> Optional[str]:
        return self.domain_synonyms.get(word.lower())

    def _add_hedging(self, text: str) -> str:
        if self.intensity < 3:
            return text
        hedges = ["may indicate","suggests that","could be attributed to","appears to be","is likely","potentially","in certain contexts"]
        sentences = sent_tokenize(text)
        new_sents = []
        for sent in sentences:
            if random.random() < 0.15 * (self.intensity/5):
                words = word_tokenize(sent)
                tagged = nltk.pos_tag(words)
                verb_positions = [i for i, (_,pos) in enumerate(tagged) if pos.startswith("VB") and i>1]
                if verb_positions:
                    idx = random.choice(verb_positions)
                else:
                    idx = min(2, len(words)-1)
                words.insert(idx, random.choice(hedges))
                new_sents.append(" ".join(words))
            else:
                new_sents.append(sent)
        return " ".join(new_sents)

    def _vary_sentence_beginnings(self, paragraph: str) -> str:
        sentences = sent_tokenize(paragraph)
        trans_list = self.transition_words.get(self.domain, self.transition_words["general"])
        new_sents = []
        for i, sent in enumerate(sentences):
            if i>0 and random.random()<0.3:
                trans = random.choice(trans_list)
                new_sents.append(f"{trans}, {sent[0].lower()}{sent[1:]}")
            else:
                new_sents.append(sent)
        return " ".join(new_sents)

    def _add_rhetorical_question(self, paragraph: str) -> str:
        questions = {
            "medical": "Could this finding reshape our understanding of disease progression?",
            "engineering": "Is this design paradigm robust enough for real-world applications?",
            "humanities": "What does this reveal about the human condition in the modern era?",
            "general": "How might these results influence future research directions?"
        }
        if random.random()<0.2:
            return paragraph + " " + questions.get(self.domain, questions["general"])
        return paragraph

    def _distort_paragraph_structure(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences)<3 or self.intensity<3:
            return text
        indices = list(range(len(sentences)))
        i = random.randint(0, len(sentences)-2)
        j = random.randint(0, len(sentences)-2)
        if i!=j:
            indices[i], indices[j] = indices[j], indices[i]
        return " ".join(sentences[k] for k in indices)

    def _protect_references(self, text: str) -> Tuple[str, Dict[str,str]]:
        replacements: Dict[str,str] = {}
        counter = 0
        def _replace_citation(match):
            nonlocal counter
            token = f"__CITATION_{counter}__"
            replacements[token] = match.group(0)
            counter += 1
            return token
        text = self.citation_pattern.sub(_replace_citation, text)
        def _replace_number(match):
            nonlocal counter
            token = f"__NUM_{counter}__"
            replacements[token] = match.group(0)
            counter += 1
            return token
        text = self.number_pattern.sub(_replace_number, text)
        return text, replacements

    def _restore_protected(self, text: str, replacements: Dict[str,str]) -> str:
        for token, orig in replacements.items():
            text = text.replace(token, orig)
        return text

    def _split_long_sentences(self, text: str, max_words: int = 45) -> str:
        sentences = sent_tokenize(text)
        new_sents = []
        for sent in sentences:
            words = word_tokenize(sent)
            if len(words) > max_words:
                mid = len(words)//2
                new_sents.append(" ".join(words[:mid]) + ";")
                new_sents.append(" ".join(words[mid:]))
            else:
                new_sents.append(sent)
        return " ".join(new_sents)

    def _remove_long_dashes(self, text: str) -> str:
        return text.replace("—", "; ").replace("–", ", ")

    # ------------------------------------------------------------------
    # المحركات الستة
    # ------------------------------------------------------------------
    def engine1_perplexity_injector(self, text: str) -> str:
        sentences = sent_tokenize(text)
        new_sentences = []
        for sent in sentences:
            words = word_tokenize(sent)
            for i, word in enumerate(words):
                if len(word)>3 and word.isalpha() and not word.startswith("__"):
                    synonym = self._get_synonym(word)
                    if synonym and random.random() < 0.3*(self.intensity/5):
                        words[i] = synonym
            new_sentences.append(" ".join(words))
        return self._add_hedging(" ".join(new_sentences))

    def engine2_burstiness_synthesizer(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences)<2:
            return text
        target_lengths = random.choices(self.reference_lengths, k=len(sentences))
        for i in range(1, len(target_lengths)):
            if abs(target_lengths[i] - target_lengths[i-1]) <= 2:
                target_lengths[i] = target_lengths[i-1] + random.choice([-3,3])
                target_lengths[i] = max(5, target_lengths[i])
        fillers = ["specifically","indeed","notably","in practice","overall"]
        new_sents = []
        for sent, target in zip(sentences, target_lengths):
            words = word_tokenize(sent)
            if len(words) < target:
                words.extend(random.choices(fillers, k=target-len(words)))
            else:
                words = words[:target]
            new_sents.append(" ".join(words))
        return " ".join(new_sents)

    def engine3_stylistic_fingerprint_forger(self, text: str) -> str:
        text = self._vary_sentence_beginnings(text)
        return self._add_rhetorical_question(text)

    def engine4_semantic_deepener(self, text: str) -> str:
        sentences = sent_tokenize(text)
        if len(sentences)<2:
            return text
        causal = self.causal_links.get(self.domain, self.causal_links["general"])
        contrast = self.contrast_links.get(self.domain, self.contrast_links["general"])
        new_sents = []
        for i, sent in enumerate(sentences):
            if i==0:
                new_sents.append(sent)
                continue
            link = random.choice(causal if random.random()<0.5 else contrast)
            if not any(sent.lower().startswith(w) for w in link.split(",")):
                new_sents.append(f"{link.capitalize()}, {sent[0].lower()}{sent[1:]}")
            else:
                new_sents.append(sent)
        if len(new_sents)>=3 and random.random()<0.2:
            indices = list(range(len(new_sents)))
            i,j = random.sample(range(1,len(new_sents)),2)
            indices[i], indices[j] = indices[j], indices[i]
            new_sents = [new_sents[k] for k in indices]
        return " ".join(new_sents)

    def engine5_watermark_distorter(self, text: str) -> str:
        return self._distort_paragraph_structure(text)

    def engine6_coherence_checker(self, original: str, modified: str) -> str:
        orig_sents = sent_tokenize(original)
        mod_sents = sent_tokenize(modified)
        if len(orig_sents) != len(mod_sents):
            return modified
        corrected = []
        for o_sent, m_sent in zip(orig_sents, mod_sents):
            o_tags = nltk.pos_tag(word_tokenize(o_sent))
            m_tags = nltk.pos_tag(word_tokenize(m_sent))
            o_verbs = [tag for _,tag in o_tags if tag.startswith("VB")]
            m_verbs = [tag for _,tag in m_tags if tag.startswith("VB")]
            if o_verbs and m_verbs and o_verbs[0] != m_verbs[0]:
                m_words = word_tokenize(m_sent)
                o_words = word_tokenize(o_sent)
                for idx, (_,tag) in enumerate(m_tags):
                    if tag.startswith("VB") and idx < len(o_words):
                        m_words[idx] = o_words[idx]
                m_sent = " ".join(m_words)
            corrected.append(m_sent)
        return " ".join(corrected)

    def semantic_lock(self, original: str, modified: str, threshold=0.92) -> bool:
        if not original.strip() or not modified.strip():
            return False
        emb_orig = self.semantic_model.encode(original, convert_to_tensor=True)
        emb_mod = self.semantic_model.encode(modified, convert_to_tensor=True)
        return util.pytorch_cos_sim(emb_orig, emb_mod).item() >= threshold

    def check_flow_extremity(self, lengths: List[int]) -> bool:
        if len(lengths)<3:
            return True
        return np.std(lengths) <= 3*np.std(self.reference_lengths)

    def run_pipeline(self, progress_callback=None) -> str:
        text = self.original_text
        protected_text, replacement_map = self._protect_references(text)

        paragraphs = [p.strip() for p in protected_text.split("\n\n") if p.strip()]
        processed = []
        context = []
        total_stages = len(paragraphs)*7
        current_stage = 0

        for para in paragraphs:
            if progress_callback:
                progress_callback("جاري التقسيم والتحليل...", current_stage/total_stages)
            para_with_context = (" ".join(context)+" "+para) if context else para
            sentences = sent_tokenize(para)
            context = sentences[-5:] if len(sentences)>=5 else sentences
            original_para = para_with_context
            modified = para_with_context

            current_stage+=1
            if progress_callback: progress_callback("جاري التعديل الإحصائي...", current_stage/total_stages)
            modified = self.engine1_perplexity_injector(modified)
            current_stage+=1
            if progress_callback: progress_callback("جاري تعديل التدافع...", current_stage/total_stages)
            modified = self.engine2_burstiness_synthesizer(modified)
            current_stage+=1
            if progress_callback: progress_callback("جاري تزوير البصمة الأسلوبية...", current_stage/total_stages)
            modified = self.engine3_stylistic_fingerprint_forger(modified)
            current_stage+=1
            if progress_callback: progress_callback("جاري تعميق الدلالة...", current_stage/total_stages)
            modified = self.engine4_semantic_deepener(modified)
            current_stage+=1
            if progress_callback: progress_callback("جاري تشويش العلامات المائية...", current_stage/total_stages)
            modified = self.engine5_watermark_distorter(modified)
            current_stage+=1
            if progress_callback: progress_callback("جاري التحقق من الاتساق...", current_stage/total_stages)
            modified = self.engine6_coherence_checker(original_para, modified)

            if not self.semantic_lock(original_para, modified):
                modified = original_para
                modified = self.engine1_perplexity_injector(modified)
                modified = self.engine4_semantic_deepener(modified)

            processed.append(modified)

        final_text = "\n\n".join(processed)
        final_text = self._restore_protected(final_text, replacement_map)
        final_text = self._remove_long_dashes(final_text)
        final_text = self._split_long_sentences(final_text)

        all_sents = sent_tokenize(final_text)
        lengths = [len(word_tokenize(s)) for s in all_sents]
        if not self.check_flow_extremity(lengths):
            final_sents = []
            for s in all_sents:
                words = word_tokenize(s)
                if len(words)>40:
                    words = words[:40]
                final_sents.append(" ".join(words))
            final_text = " ".join(final_sents)

        return final_text


# =============================================================================
# دوال مساعدة للواجهة
# =============================================================================
def create_word_document(text: str, title="DeepClean Studio Output") -> BytesIO:
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for para in text.split('\n'):
        if para.strip():
            p = doc.add_paragraph(para.strip())
            p.style.font.size = Pt(12)
            p.paragraph_format.space_after = Pt(6)
        else:
            doc.add_paragraph()
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def word_level_diff(original: str, modified: str) -> str:
    orig_words = original.split()
    mod_words = modified.split()
    diff = difflib.SequenceMatcher(None, orig_words, mod_words)
    html_parts = []
    for tag, i1, i2, j1, j2 in diff.get_opcodes():
        if tag == 'equal':
            html_parts.append(' '.join(orig_words[i1:i2]))
        elif tag == 'replace':
            html_parts.append(f"<span style='color:red;text-decoration:line-through;'>{' '.join(orig_words[i1:i2])}</span> ")
            html_parts.append(f"<span style='color:green;font-weight:bold;'>{' '.join(mod_words[j1:j2])}</span> ")
        elif tag == 'delete':
            html_parts.append(f"<span style='color:red;text-decoration:line-through;'>{' '.join(orig_words[i1:i2])}</span> ")
        elif tag == 'insert':
            html_parts.append(f"<span style='color:green;font-weight:bold;'>{' '.join(mod_words[j1:j2])}</span> ")
    return ' '.join(html_parts)

def compute_metrics(original: str, enhanced: str):
    try:
        words_orig = word_tokenize(original.lower())
        words_enh = word_tokenize(enhanced.lower())
        freq_orig = nltk.FreqDist(words_orig)
        log_prob_sum = sum(math.log(max(freq_orig.freq(w),1e-10)) for w in words_enh if w.isalpha())
        count = sum(1 for w in words_enh if w.isalpha())
        perplexity = math.exp(-log_prob_sum/count) if count else 0.0
        lengths = [len(word_tokenize(s)) for s in sent_tokenize(enhanced)]
        burstiness = float(np.std(lengths)/np.mean(lengths)) if np.mean(lengths)>0 else 0.0
        tokens_enh = [w.lower() for w in words_enh if w.isalpha()]
        ttr = len(set(tokens_enh))/len(tokens_enh) if tokens_enh else 0.0
        return perplexity, burstiness, ttr
    except Exception:
        return 0.0, 0.0, 0.0

# =============================================================================
# دالة المعالجة الأساسية (callback)
# =============================================================================
def process_text_callback():
    """تُستدعى عند النقر على زر المعالجة. تخزن النتائج في session_state."""
    if not st.session_state.get("text_input", ""):
        return

    engine = DeepCleanEngine(
        domain=st.session_state.get("domain", "general"),
        intensity=st.session_state.get("intensity", 3),
        text=st.session_state["text_input"]
    )

    # لا نستخدم update_progress هنا لأنها تتطلب session_state
    enhanced = engine.run_pipeline()

    st.session_state.enhanced_text = enhanced
    st.session_state.perplexity, st.session_state.burstiness, st.session_state.ttr = \
        compute_metrics(st.session_state["text_input"], enhanced)
    st.session_state.processing_done = True


# =============================================================================
# واجهة المستخدم – Streamlit
# =============================================================================
st.set_page_config(page_title="DeepClean Studio", layout="wide")
st.title("🛡️ DeepClean Studio")
st.markdown("تحويل النصوص الأكاديمية من نمط الذكاء الاصطناعي إلى طابع بشري خبير")

# تهيئة session_state
defaults = {
    "perplexity": 0.0, "burstiness": 0.0, "ttr": 0.0,
    "enhanced_text": "", "processing_done": False,
    "text_input": "", "domain": "medical", "intensity": 3
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ------------------------- الشريط الجانبي -------------------------
with st.sidebar:
    st.header("⚙️ الإعدادات")
    input_option = st.radio("مصدر النص:", ("رفع ملف", "لصق نص"), key="source_radio")

    uploaded_file = None
    text_input = ""

    if input_option == "رفع ملف":
        uploaded_file = st.file_uploader("اختر ملفًا (txt, docx, pdf)", type=["txt","docx","pdf"], key="file_uploader")
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".txt"):
                    text_input = uploaded_file.read().decode("utf-8")
                elif uploaded_file.name.endswith(".docx"):
                    text_input = docx2txt.process(uploaded_file)
                elif uploaded_file.name.endswith(".pdf"):
                    pdf_reader = pypdf.PdfReader(uploaded_file)
                    text_input = "".join(page.extract_text() for page in pdf_reader.pages)
            except Exception as e:
                st.error(f"خطأ في قراءة الملف: {e}")
    else:
        text_input = st.text_area("ألصق النص الأكاديمي هنا:", height=200, key="paste_area")

    # حفظ النص المدخل في session_state
    if text_input:
        st.session_state.text_input = text_input

    intensity = st.slider("قوة التحويل", 1, 5, 3, help="1=دقيق ومحافظ، 5=تحول إبداعي", key="intensity")
    domain = st.selectbox("المجال الأكاديمي:", ("medical","engineering","humanities","general"),
                          format_func=lambda x: {"medical":"طبي","engineering":"هندسي","humanities":"علوم إنسانية","general":"عام"}[x],
                          key="domain")

    # استخدام on_click لتجنب st.rerun()
    process_btn = st.button(
        "🛡️ بدء التحويل الآمن",
        type="primary",
        use_container_width=True,
        on_click=process_text_callback
    )

    show_changes = st.checkbox("عرض التغييرات للمراجعة البشرية")

    st.markdown("---")
    st.subheader("📊 مؤشرات الفحص الذاتي (محلية)")
    st.metric("Perplexity (تقريبي)", f"{st.session_state.perplexity:.2f}")
    st.metric("Burstiness Score", f"{st.session_state.burstiness:.2f}")
    st.metric("TTR Ratio", f"{st.session_state.ttr:.2f}")
    st.warning("هذه مؤشرات محلية فقط ولا تضمن اجتياز أي كاشف خارجي.", icon="⚠️")

    # زر تحميل النتيجة كـ Word
    if st.session_state.enhanced_text:
        word_file = create_word_document(st.session_state.enhanced_text)
        st.download_button("📥 تنزيل النص المحسَّن (Word)", data=word_file,
                           file_name="deepclean_output.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ------------------------- المنطقة الرئيسية -------------------------
col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 النص الأصلي")
    if st.session_state.text_input:
        st.text_area("", st.session_state.text_input, height=400, key="orig_text_display")
        st.caption(f"عدد الكلمات: {len(st.session_state.text_input.split())}")

with col2:
    st.subheader("🧬 النص المحسّن")
    if st.session_state.enhanced_text:
        st.text_area("", st.session_state.enhanced_text, height=400, key="enh_text_display")
        st.caption(f"عدد الكلمات: {len(st.session_state.enhanced_text.split())}")

# ------------------------- عرض التغييرات -------------------------
if show_changes and st.session_state.enhanced_text and st.session_state.text_input:
    st.markdown("---")
    st.subheader("🔍 تفاصيل التغييرات (مقارنة كلمة بكلمة)")

    diff_html = word_level_diff(st.session_state.text_input, st.session_state.enhanced_text)
    st.markdown(
        f"""<div style='background-color:#f9f9f9; padding:15px; border-radius:8px; line-height:1.6;'>
        {diff_html}</div>""",
        unsafe_allow_html=True
    )

    orig_words_set = set(word_tokenize(st.session_state.text_input.lower()))
    enh_words_set = set(word_tokenize(st.session_state.enhanced_text.lower()))
    removed = [w for w in orig_words_set if w not in enh_words_set and w.isalpha()]
    added = [w for w in enh_words_set if w not in orig_words_set and w.isalpha()]
    if removed or added:
        df_changes = pd.DataFrame({
            "النوع": ["محذوف"]*len(removed) + ["مضاف"]*len(added),
            "الكلمة": removed + added
        })
        st.dataframe(df_changes, use_container_width=True)
    else:
        st.info("لا توجد تغييرات جوهرية على مستوى الكلمات.")

st.markdown("---")
st.markdown("DeepClean Studio © 2025 – للأغراض التعليمية والبحثية فقط")
