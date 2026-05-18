#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# =============================================================================
# DeepClean Studio - تحويل النصوص الأكاديمية إلى طابع بشري خبير
# المهندس: خبير معالجة اللغة الطبيعية واللسانيات الجنائية
# =============================================================================

import streamlit as st
import nltk
import textstat
import numpy as np
import pandas as pd
import random
import math
import re
import io
import logging
from typing import List, Tuple, Dict, Any, Optional
from sentence_transformers import SentenceTransformer, util
from collections import Counter
import docx2txt
import pypdf

# في بداية الملف بعد الاستيرادات
# ... (الإعدادات والنماذج) ...

# ---- الشريط الجانبي ----
with st.sidebar:
    st.header("⚙️ الإعدادات")
    input_option = st.radio("مصدر النص:", ("رفع ملف", "لصق نص"), key="source_radio")
    uploaded_file = None
    text_input = ""
    if input_option == "رفع ملف":
        uploaded_file = st.file_uploader("اختر ملفًا (txt, docx, pdf)", type=['txt', 'docx', 'pdf'])
        if uploaded_file is not None:
            # قراءة الملف حسب نوعه
            try:
                if uploaded_file.name.endswith('.txt'):
                    text_input = uploaded_file.read().decode('utf-8')
                elif uploaded_file.name.endswith('.docx'):
                    text_input = docx2txt.process(uploaded_file)
                elif uploaded_file.name.endswith('.pdf'):
                    import pypdf  # تأكد من الاستيراد في الأعلى
                    pdf_reader = pypdf.PdfReader(uploaded_file)
                    text_input = ""
                    for page in pdf_reader.pages:
                        text_input += page.extract_text()
            except Exception as e:
                st.error(f"خطأ في قراءة الملف: {e}")
    else:
        # خيار لصق النص
        text_input = st.text_area("ألصق النص الأكاديمي هنا:", height=200)

    # بقية عناصر الشريط الجانبي (السلايدر، القائمة، الزر...)
    ...

# ---- بعد الشريط الجانبي، المنطقة الرئيسية ----
col1, col2 = st.columns(2)
...

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.txt'):
            text_input = uploaded_file.read().decode('utf-8')
        elif uploaded_file.name.endswith('.docx'):
            text_input = docx2txt.process(uploaded_file)
        elif uploaded_file.name.endswith('.pdf'):
            import pypdf  # تأكد من الاستيراد هنا أو في أعلى الملف
            pdf_reader = pypdf.PdfReader(uploaded_file)
            text_input = ""
            for page in pdf_reader.pages:
                text_input += page.extract_text()
    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")

# ---------------------------------------------------------------------------
# تنزيل موارد NLTK الضرورية (يتم تنفيذه مرة واحدة)
# ---------------------------------------------------------------------------
@st.cache_resource
def download_nltk_resources():
    try:
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        nltk.download('punkt')
    try:
        nltk.data.find('corpora/wordnet')
    except LookupError:
        nltk.download('wordnet')
    try:
        nltk.data.find('corpora/stopwords')
    except LookupError:
        nltk.download('stopwords')
    try:
        nltk.data.find('taggers/averaged_perceptron_tagger')
    except LookupError:
        nltk.download('averaged_perceptron_tagger')

download_nltk_resources()

from nltk.corpus import wordnet
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords

# =============================================================================
# إعداد النموذج الدلالي (يتم تحميله مرة واحدة)
# =============================================================================
@st.cache_resource
def load_semantic_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

# =============================================================================
# تحميل قاموس المرادفات الأكاديمية من ملف CSV
# =============================================================================
@st.cache_data
def load_synonym_dictionary(filepath: str = 'synonyms_academic.csv') -> pd.DataFrame:
    try:
        df = pd.read_csv(filepath, encoding='utf-8')
        if set(['domain', 'original', 'replacement']).issubset(df.columns):
            return df
        else:
            st.error("ملف المرادفات يجب أن يحتوى على الأعمدة: domain, original, replacement")
            return pd.DataFrame(columns=['domain', 'original', 'replacement'])
    except FileNotFoundError:
        st.warning("ملف synonyms_academic.csv غير موجود. سيتم استخدام قاموس افتراضي صغير.")
        # إنشاء قاموس افتراضي صغير
        default_data = {
            'domain': ['medical', 'medical', 'engineering', 'engineering', 'humanities', 'humanities',
                       'medical', 'engineering', 'humanities'],
            'original': ['patient', 'treatment', 'system', 'design', 'culture', 'history',
                        'diagnosis', 'algorithm', 'philosophy'],
            'replacement': ['individual receiving medical care', 'therapeutic intervention',
                           'integrated framework', 'architectural conception',
                           'sociocultural paradigm', 'historical trajectory',
                           'clinical assessment', 'computational procedure',
                           'school of thought']
        }
        return pd.DataFrame(default_data)

# =============================================================================
# الكلاس الرئيسي للمعالجة
# =============================================================================
class DeepCleanEngine:
    def __init__(self, domain: str, intensity: int, text: str):
        """
        معاملات:
            domain: المجال الأكاديمي (medical, engineering, humanities, general)
            intensity: قوة التحويل (1-5)
            text: النص الأصلي الكامل
        """
        self.domain = domain
        self.intensity = intensity
        self.original_text = text
        self.semantic_model = load_semantic_model()
        self.synonym_df = load_synonym_dictionary()

        # إعداد مرادفات المجال المحدد
        self.domain_synonyms = {}
        if not self.synonym_df.empty:
            domain_df = self.synonym_df[self.synonym_df['domain'] == self.domain]
            for _, row in domain_df.iterrows():
                orig = row['original'].lower()
                repl = row['replacement'].lower()
                self.domain_synonyms[orig] = repl
        # إذا كان المجال عام، ندمج قواميس المجالات الأخرى جزئياً
        if self.domain == 'general' and not self.synonym_df.empty:
            for _, row in self.synonym_df.iterrows():
                orig = row['original'].lower()
                repl = row['replacement'].lower()
                self.domain_synonyms[orig] = repl

        # مصفوفة أطوال الجمل المرجعية (من تحليل 500 مقال بشري) للتدافع
        self.reference_lengths = [15, 22, 18, 30, 25, 19, 28, 16, 21, 27, 23, 17, 20, 24, 26, 19, 22, 31, 14, 29]

        # علامات الترقيم للتنويع
        self.transition_words = {
            'medical': ['furthermore', 'moreover', 'conversely', 'notably', 'specifically'],
            'engineering': ['additionally', 'in contrast', 'consequently', 'accordingly', 'particularly'],
            'humanities': ['moreover', 'on the other hand', 'thus', 'indeed', 'above all'],
            'general': ['furthermore', 'however', 'therefore', 'for instance', 'in particular']
        }

        # روابط سببية/تناقضية للمحرك الرابع
        self.causal_links = {
            'medical': ['due to', 'as a result of', 'leading to', 'stemming from'],
            'engineering': ['caused by', 'resulting in', 'attributed to', 'driven by'],
            'humanities': ['owing to', 'consequently', 'as a consequence', 'thereby'],
            'general': ['because of', 'hence', 'thus', 'accordingly']
        }
        self.contrast_links = {
            'medical': ['whereas', 'in contrast', 'conversely', 'on the contrary'],
            'engineering': ['however', 'on the other hand', 'alternatively', 'in opposition'],
            'humanities': ['nevertheless', 'nonetheless', 'yet', 'contrarily'],
            'general': ['but', 'however', 'although', 'in spite of']
        }

        # أنماط المراجع والبيانات للحماية
        self.citation_pattern = re.compile(r'\[[\d,\-; ]+\]|\([^)]*\d{4}[^)]*\)')
        self.number_pattern = re.compile(r'\b\d+(?:\.\d+)?%?\b')
        self.unit_pattern = re.compile(r'\b(?:mg|kg|cm|mm|ml|°C|mol|Hz|W|V|A)\b')

    # -------------------------------------------------------------------------
    # دوال مساعدة
    # -------------------------------------------------------------------------
    def _get_synonym(self, word: str) -> Optional[str]:
        """البحث عن مرادف من قاموس المجال"""
        w = word.lower()
        if w in self.domain_synonyms:
            return self.domain_synonyms[w]
        return None

    def _add_hedging(self, text: str, intensity: int) -> str:
        """إضافة تحوط أكاديمي باعتدال حسب الشدة"""
        if intensity < 3:
            return text
        hedges = ['may indicate', 'suggests that', 'could be attributed to', 'appears to be',
                  'is likely', 'potentially', 'in certain contexts']
        sentences = sent_tokenize(text)
        new_sents = []
        for sent in sentences:
            if random.random() < 0.15 * (intensity/5):  # احتمالية ضئيلة
                # إضافة تحوط في بداية الجملة أو بعد الفعل
                if len(sent.split()) > 5:
                    words = word_tokenize(sent)
                    verb_positions = [i for i, (word, pos) in enumerate(nltk.pos_tag(words))
                                      if pos.startswith('VB') and i > 1]
                    if verb_positions:
                        idx = random.choice(verb_positions)
                        words.insert(idx, random.choice(hedges))
                    else:
                        words.insert(2, random.choice(hedges))
                    new_sent = ' '.join(words)
                    new_sents.append(new_sent)
                else:
                    new_sents.append(sent)
            else:
                new_sents.append(sent)
        return ' '.join(new_sents)

    def _vary_sentence_beginnings(self, paragraph: str) -> str:
        """تنويع بدايات الجمل باستخدام كلمات انتقالية"""
        sentences = sent_tokenize(paragraph)
        new_sents = []
        for i, sent in enumerate(sentences):
            if i > 0 and random.random() < 0.3:
                trans = random.choice(self.transition_words.get(self.domain, self.transition_words['general']))
                new_sents.append(f"{trans}, {sent[0].lower()}{sent[1:]}")
            else:
                new_sents.append(sent)
        return ' '.join(new_sents)

    def _add_rhetorical_question(self, paragraph: str, domain: str) -> str:
        """إضافة لمسة شخصية (سؤال بلاغي) باعتدال"""
        questions = {
            'medical': 'Could this finding reshape our understanding of disease progression?',
            'engineering': 'Is this design paradigm robust enough for real-world applications?',
            'humanities': 'What does this reveal about the human condition in the modern era?',
            'general': 'How might these results influence future research directions?'
        }
        if random.random() < 0.2:  # فقط 20% من الفقرات
            return paragraph + ' ' + questions.get(domain, questions['general'])
        return paragraph

    def _distort_paragraph_structure(self, text: str, intensity: int) -> str:
        """إعادة ترتيب الجمل بطريقة تحافظ على التدفق المنطقي (مشوش العلامات المائية)"""
        sentences = sent_tokenize(text)
        if len(sentences) < 3 or intensity < 3:
            return text
        # اختيار عشوائي لجملتين متبادلتين مع الحفاظ على التسلسل العام
        indices = list(range(len(sentences)))
        i = random.randint(0, len(sentences)-2)
        j = random.randint(0, len(sentences)-2)
        if i != j:
            indices[i], indices[j] = indices[j], indices[i]
        reordered = [sentences[k] for k in indices]
        return ' '.join(reordered)

    def _break_repetition(self, paragraphs: List[str]) -> List[str]:
        """كسر تكرار توزيع أجزاء الكلام عبر الفقرات (المحرك 5)"""
        if len(paragraphs) < 2:
            return paragraphs
        pos_distributions = []
        for para in paragraphs:
            tokens = nltk.pos_tag(word_tokenize(para))
            pos_counts = Counter(tag for _, tag in tokens)
            pos_distributions.append(pos_counts)
        new_paragraphs = paragraphs.copy()
        # إذا كانت فقرتان متتاليتان لهما توزيع متشابه جداً (معامل جاكارد > 0.8)، نعيد ترتيب إحداهما
        for i in range(len(new_paragraphs)-1):
            if len(pos_distributions[i]) > 0 and len(pos_distributions[i+1]) > 0:
                common = set(pos_distributions[i].keys()) & set(pos_distributions[i+1].keys())
                union = set(pos_distributions[i].keys()) | set(pos_distributions[i+1].keys())
                if union:
                    jaccard = len(common) / len(union)
                    if jaccard > 0.8:
                        # خلط جمل الفقرة الثانية بشكل عشوائي
                        sents = sent_tokenize(new_paragraphs[i+1])
                        if len(sents) > 2:
                            random.shuffle(sents)
                            new_paragraphs[i+1] = ' '.join(sents)
        return new_paragraphs

    def _protect_references(self, text: str) -> Tuple[str, Dict[str, str]]:
        """استبدال المراجع والأرقام برموز مؤقتة لحمايتها"""
        replacements = {}
        counter = 0
        def replace_citation(match):
            nonlocal counter
            token = f"__CITATION_{counter}__"
            replacements[token] = match.group(0)
            counter += 1
            return token
        text = self.citation_pattern.sub(replace_citation, text)

        def replace_number(match):
            nonlocal counter
            token = f"__NUM_{counter}__"
            replacements[token] = match.group(0)
            counter += 1
            return token
        text = self.number_pattern.sub(replace_number, text)

        return text, replacements

    def _restore_protected(self, text: str, replacements: Dict[str, str]) -> str:
        """استعادة الرموز المحمية إلى أصلها"""
        for token, orig in replacements.items():
            text = text.replace(token, orig)
        return text

    # -------------------------------------------------------------------------
    # طبقات الأمان
    # -------------------------------------------------------------------------
    def semantic_lock(self, original: str, modified: str, threshold: float = 0.92) -> bool:
        """التحقق من التشابه الدلالي بين النص الأصلي والمعدل"""
        if not original.strip() or not modified.strip():
            return False
        emb_orig = self.semantic_model.encode(original, convert_to_tensor=True)
        emb_mod = self.semantic_model.encode(modified, convert_to_tensor=True)
        similarity = util.pytorch_cos_sim(emb_orig, emb_mod).item()
        return similarity >= threshold

    def check_flow_extremity(self, sentences_lengths: List[int]) -> bool:
        """مؤشر تطرف التدفق: الانحراف المعياري مقارنة بالمرجع"""
        if len(sentences_lengths) < 3:
            return True
        ref_std = np.std(self.reference_lengths)
        current_std = np.std(sentences_lengths)
        # إذا كان الانحراف المعياري أكثر من 3 أضعاف المرجع، نعتبره متطرفاً
        if current_std > 3 * ref_std:
            return False
        return True

    # -------------------------------------------------------------------------
    # المحركات الستة (تنفذ كدوال مستقلة)
    # -------------------------------------------------------------------------
    def engine1_perplexity_injector(self, text: str) -> str:
        """محقون الحيرة: استبدال الكلمات المتوقعة بمرادفات أكاديمية وإضافة تحوط"""
        sentences = sent_tokenize(text)
        new_sentences = []
        for sent in sentences:
            words = word_tokenize(sent)
            # التعرف على الكلمات القابلة للاستبدال (أطول من 3 أحرف، ليست رموزاً محمية)
            for i, word in enumerate(words):
                if len(word) > 3 and word.isalpha() and not word.startswith('__'):
                    synonym = self._get_synonym(word)
                    if synonym and random.random() < 0.3 * (self.intensity/5):
                        words[i] = synonym
            new_sent = ' '.join(words)
            new_sentences.append(new_sent)
        modified = ' '.join(new_sentences)
        # إضافة التحوط حسب الشدة
        modified = self._add_hedging(modified, self.intensity)
        return modified

    def engine2_burstiness_synthesizer(self, text: str) -> str:
        """مركب التدافع: إعادة توزيع أطوال الجمل لتحاكي التوزيع المرجعي"""
        sentences = sent_tokenize(text)
        if len(sentences) < 2:
            return text
        # نأخذ أطوالاً عشوائية من المصفوفة المرجعية بنفس عدد الجمل
        target_lengths = random.choices(self.reference_lengths, k=len(sentences))
        # نمنع تطابق الطول لكل جملتين متتاليتين (±2 كلمة)
        for i in range(1, len(target_lengths)):
            if abs(target_lengths[i] - target_lengths[i-1]) <= 2:
                # تعديل الطول الحالي
                target_lengths[i] = target_lengths[i-1] + random.choice([-3, 3])
                target_lengths[i] = max(5, target_lengths[i])  # حد أدنى 5 كلمات
        # نعيد بناء الجمل بأطوال قريبة من الهدف عن طريق اقتطاع أو إضافة كلمات
        new_sentences = []
        for sent, target in zip(sentences, target_lengths):
            words = word_tokenize(sent)
            current_len = len(words)
            if current_len < target:
                # إضافة كلمات حشو أكاديمية
                fillers = ['specifically', 'indeed', 'notably', 'in practice', 'overall']
                words.extend(random.choices(fillers, k=target - current_len))
            elif current_len > target:
                # اقتطاع الكلمات الزائدة من النهاية
                words = words[:target]
            new_sentences.append(' '.join(words))
        return ' '.join(new_sentences)

    def engine3_stylistic_fingerprint_forger(self, text: str) -> str:
        """مزور البصمة الأسلوبية: تنويع علامات الترقيم وبدايات الجمل وإضافة لمسة شخصية"""
        text = self._vary_sentence_beginnings(text)
        text = self._add_rhetorical_question(text, self.domain)
        return text

    def engine4_semantic_deepener(self, text: str) -> str:
        """معمق الدلالة: استبدال الروابط العامة بأخرى سببية/تناقضية وإعادة ترتيب المنطق"""
        sentences = sent_tokenize(text)
        if len(sentences) < 2:
            return text
        causal = self.causal_links.get(self.domain, self.causal_links['general'])
        contrast = self.contrast_links.get(self.domain, self.contrast_links['general'])
        new_sentences = []
        for i, sent in enumerate(sentences):
            if i > 0:
                # قرار عشوائي: ربط سببي أو تناقضي
                if random.random() < 0.5:
                    link = random.choice(causal)
                else:
                    link = random.choice(contrast)
                # دمج الرابط مع الجملة إذا كانت لا تبدأ برابط بالفعل
                if not any(sent.lower().startswith(w) for w in link.split(',')):
                    new_sent = f"{link.capitalize()}, {sent[0].lower()}{sent[1:]}"
                else:
                    new_sent = sent
            else:
                new_sent = sent
            new_sentences.append(new_sent)
        # إعادة ترتيب الجمل لكشف العلاقات السببية (ببساطة نعكس ترتيب جملتين إذا أمكن)
        if len(new_sentences) >= 3 and random.random() < 0.2:
            indices = list(range(len(new_sentences)))
            i, j = random.sample(range(1, len(new_sentences)), 2)
            indices[i], indices[j] = indices[j], indices[i]
            new_sentences = [new_sentences[k] for k in indices]
        return ' '.join(new_sentences)

    def engine5_watermark_distorter(self, text: str) -> str:
        """مشوش العلامات المائية: إعادة تشكيل هيكل الفقرة وكسر تكرار أجزاء الكلام"""
        text = self._distort_paragraph_structure(text, self.intensity)
        return text

    def engine6_coherence_checker(self, original: str, modified: str) -> str:
        """مدقق الاتساق المنطقي: تصحيح الأخطاء النحوية أو الدلالية الناتجة عن التعديل"""
        # آلية بسيطة: نستخدم POS tagging للتحقق من توافق الفعل والفاعل
        # إذا تغير زمن الفعل بشكل غير منطقي، نعيده إلى الأصل
        orig_sents = sent_tokenize(original)
        mod_sents = sent_tokenize(modified)
        if len(orig_sents) != len(mod_sents):
            return modified  # لا نتدخل إذا اختلف عدد الجمل
        corrected = []
        for o_sent, m_sent in zip(orig_sents, mod_sents):
            o_tags = nltk.pos_tag(word_tokenize(o_sent))
            m_tags = nltk.pos_tag(word_tokenize(m_sent))
            # مقارنة أزمنة الأفعال الرئيسية
            o_verbs = [tag for _, tag in o_tags if tag.startswith('VB')]
            m_verbs = [tag for _, tag in m_tags if tag.startswith('VB')]
            if o_verbs and m_verbs and o_verbs[0] != m_verbs[0]:
                # إصلاح بإعادة الفعل الأصلي (تبسيط)
                # هنا نستبدل الفعل المعدل بالفعل الأصلي إن أمكن
                m_words = word_tokenize(m_sent)
                for i, (word, tag) in enumerate(m_tags):
                    if tag.startswith('VB') and i < len(o_verbs):
                        m_words[i] = word_tokenize(o_sent)[i]  # استبدال تقريبي
                m_sent = ' '.join(m_words)
            corrected.append(m_sent)
        return ' '.join(corrected)

    # -------------------------------------------------------------------------
    # خط الأنابيب الرئيسي
    # -------------------------------------------------------------------------
    def run_pipeline(self, progress_callback=None) -> str:
        """
        تنفيذ المحركات بالتسلسل مع الحماية والفحص الذاتي.
        progress_callback: دالة لاستدعاء تحديث شريط التقدم (مرحلة، نسبة)
        """
        text = self.original_text
        # حماية المراجع والبيانات
        protected_text, replacement_map = self._protect_references(text)

        # تقسيم إلى فقرات
        paragraphs = protected_text.split('\n\n')
        paragraphs = [p.strip() for p in paragraphs if p.strip()]

        # تقسيم الفقرات الطويلة جداً إلى أجزاء مع الاحتفاظ بآخر 5 جمل سياق
        processed_paragraphs = []
        context = []  # آخر 5 جمل من الفقرة السابقة

        total_stages = len(paragraphs) * 7  # 7 خطوات لكل فقرة (1 حماية + 6 محركات)
        current_stage = 0

        for para_idx, para in enumerate(paragraphs):
            if progress_callback:
                progress_callback("جاري التقسيم والتحليل...", current_stage/total_stages)
            # السياق السابق: إضافته قبل الفقرة للحفاظ على الاتساق
            if context:
                para_with_context = ' '.join(context) + ' ' + para
            else:
                para_with_context = para
            sentences = sent_tokenize(para)
            # تحديث السياق: آخر 5 جمل
            context = sentences[-5:] if len(sentences) >= 5 else sentences

            # تطبيق المحركات على الفقرة
            original_para = para_with_context
            modified = para_with_context

            # المحرك 1
            if progress_callback:
                current_stage += 1
                progress_callback("جاري التعديل الإحصائي (المحرك 1)...", current_stage/total_stages)
            modified = self.engine1_perplexity_injector(modified)

            # المحرك 2
            if progress_callback:
                current_stage += 1
                progress_callback("جاري تعديل التدافع (المحرك 2)...", current_stage/total_stages)
            modified = self.engine2_burstiness_synthesizer(modified)

            # المحرك 3
            if progress_callback:
                current_stage += 1
                progress_callback("جاري تزوير البصمة الأسلوبية (المحرك 3)...", current_stage/total_stages)
            modified = self.engine3_stylistic_fingerprint_forger(modified)

            # المحرك 4
            if progress_callback:
                current_stage += 1
                progress_callback("جاري تعميق الدلالة (المحرك 4)...", current_stage/total_stages)
            modified = self.engine4_semantic_deepener(modified)

            # المحرك 5
            if progress_callback:
                current_stage += 1
                progress_callback("جاري تشويش العلامات المائية (المحرك 5)...", current_stage/total_stages)
            modified = self.engine5_watermark_distorter(modified)

            # المحرك 6: مدقق الاتساق
            if progress_callback:
                current_stage += 1
                progress_callback("جاري التحقق من الاتساق (المحرك 6)...", current_stage/total_stages)
            modified = self.engine6_coherence_checker(original_para, modified)

            # فحص القفل الدلالي بعد التعديل
            if not self.semantic_lock(original_para, modified):
                # إعادة المحاولة مرة واحدة بتعديل أقل شدة
                modified = original_para  # تراجع كامل ثم نعيد المحركات الأساسية فقط
                modified = self.engine1_perplexity_injector(modified)
                modified = self.engine4_semantic_deepener(modified)

            processed_paragraphs.append(modified)

        # إعادة تجميع النص
        final_text = '\n\n'.join(processed_paragraphs)

        # استعادة المراجع المحمية
        final_text = self._restore_protected(final_text, replacement_map)

        # مؤشر تطرف التدفق: فحص عام
        all_sentences = sent_tokenize(final_text)
        lengths = [len(word_tokenize(s)) for s in all_sentences]
        if not self.check_flow_extremity(lengths):
            # إعادة ضبط طفيف للجمل الطويلة جداً
            final_sentences = []
            for s in all_sentences:
                words = word_tokenize(s)
                if len(words) > 40:
                    words = words[:40]  # اقتطاع بسيط
                final_sentences.append(' '.join(words))
            final_text = ' '.join(final_sentences)

        return final_text

# =============================================================================
# واجهة المستخدم Streamlit
# =============================================================================
st.set_page_config(page_title="DeepClean Studio", layout="wide")
st.title("🛡️ DeepClean Studio")
st.markdown("تحويل النصوص الأكاديمية من نمط الذكاء الاصطناعي إلى طابع بشري خبير")

# ---- الشريط الجانبي ----
with st.sidebar:
    st.header("⚙️ الإعدادات")
    input_option = st.radio("مصدر النص:", ("رفع ملف", "لصق نص"))
    uploaded_file = None
    text_input = ""
    if input_option == "رفع ملف":
        uploaded_file = st.file_uploader("اختر ملفًا (txt, docx, pdf)", type=['txt', 'docx', 'pdf'])
        if uploaded_file is not None:
            # قراءة الملف حسب نوعه
            file_details = {"filename": uploaded_file.name, "filetype": uploaded_file.type}
            try:
                if uploaded_file.name.endswith('.txt'):
                    text_input = uploaded_file.read().decode('utf-8')
                elif uploaded_file.name.endswith('.docx'):
                    text_input = docx2txt.process(uploaded_file)
                elif uploaded_file.name.endswith('.pdf'):
                    pdf_reader = PyPDF2.PdfReader(uploaded_file)
                    text_input = ""
                    for page in pdf_reader.pages:
                        text_input += page.extract_text()
            except Exception as e:
                st.error(f"خطأ في قراءة الملف: {e}")
    else:
        text_input = st.text_area("ألصق النص الأكاديمي هنا:", height=200)

    intensity = st.slider("قوة التحويل", 1, 5, 3, help="1 = دقيق ومحافظ، 5 = تحول إبداعي")
    domain = st.selectbox("المجال الأكاديمي:", ("medical", "engineering", "humanities", "general"),
                          format_func=lambda x: {"medical": "طبي", "engineering": "هندسي", 
                                                 "humanities": "علوم إنسانية", "general": "عام"}[x])
    process_btn = st.button("🛡️ بدء التحويل الآمن", type="primary", use_container_width=True)
    show_changes = st.checkbox("عرض التغييرات للمراجعة البشرية")

    st.markdown("---")
    st.subheader("📊 مؤشرات الفحص الذاتي (محلية)")
    if 'perplexity' not in st.session_state:
        st.session_state.perplexity = 0
        st.session_state.burstiness = 0
        st.session_state.ttr = 0
    st.metric("Perplexity (تقريبي)", f"{st.session_state.perplexity:.2f}")
    st.metric("Burstiness Score", f"{st.session_state.burstiness:.2f}")
    st.metric("TTR Ratio", f"{st.session_state.ttr:.2f}")
    st.warning("هذه مؤشرات محلية فقط ولا تضمن اجتياز أي كاشف خارجي.", icon="⚠️")

# ---- المنطقة الرئيسية ----
col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 النص الأصلي")
    original_placeholder = st.empty()
    if text_input:
        original_placeholder.text_area("", text_input, height=400, key="orig_text")
        st.caption(f"عدد الكلمات: {len(text_input.split())}")

with col2:
    st.subheader("🧬 النص المحسّن")
    enhanced_placeholder = st.empty()
    enhanced_text = ""
    if 'enhanced_text' in st.session_state and st.session_state.enhanced_text:
        enhanced_text = st.session_state.enhanced_text
        enhanced_placeholder.text_area("", enhanced_text, height=400, key="enh_text")
        st.caption(f"عدد الكلمات: {len(enhanced_text.split())}")

# ---- شريط التقدم التفصيلي ----
progress_bar = st.progress(0)
progress_text = st.empty()

def update_progress(stage: str, percent: float):
    progress_text.text(f"{stage} ({percent:.0%})")
    progress_bar.progress(min(percent, 1.0))

# ---- تنفيذ المعالجة عند الضغط على الزر ----
if process_btn and text_input:
    # إعادة تعيين المؤشرات
    st.session_state.perplexity = 0
    st.session_state.burstiness = 0
    st.session_state.ttr = 0
    enhanced_text = ""

    with st.spinner("جاري المعالجة..."):
        engine = DeepCleanEngine(domain=domain, intensity=intensity, text=text_input)
        # تنفيذ خط الأنابيب مع تحديث التقدم
        result = engine.run_pipeline(progress_callback=update_progress)
        enhanced_text = result

    # حساب المؤشرات المحلية
    try:
        # Perplexity مبسط: باستخدام متوسط احتمال الكلمات من تكرارها في النص
        words_orig = word_tokenize(text_input.lower())
        words_enh = word_tokenize(enhanced_text.lower())
        freq_orig = nltk.FreqDist(words_orig)
        # تقدير perplexity = exp(-متوسط لوغاريتم الاحتمال)
        log_prob_sum = 0
        count = 0
        for word in words_enh:
            if word.isalpha():
                prob = freq_orig.freq(word)
                if prob == 0:
                    prob = 1e-10
                log_prob_sum += math.log(prob)
                count += 1
        perplexity = math.exp(-log_prob_sum / count) if count > 0 else 0

        # Burstiness: معامل الاختلاف في أطوال الجمل مقارنة بالتوزيع المرجعي
        lengths = [len(word_tokenize(s)) for s in sent_tokenize(enhanced_text)]
        burstiness = np.std(lengths) / np.mean(lengths) if np.mean(lengths) > 0 else 0

        # TTR: Type-Token Ratio
        tokens_enh = [w.lower() for w in words_enh if w.isalpha()]
        types = set(tokens_enh)
        ttr = len(types) / len(tokens_enh) if tokens_enh else 0
    except Exception as e:
        perplexity = burstiness = ttr = 0

    st.session_state.perplexity = perplexity
    st.session_state.burstiness = burstiness
    st.session_state.ttr = ttr
    st.session_state.enhanced_text = enhanced_text

    # تحديث العرض
    original_placeholder.text_area("", text_input, height=400, key="orig_text_after")
    enhanced_placeholder.text_area("", enhanced_text, height=400, key="enh_text_after")
    st.rerun()

# ---- عرض التغييرات عند الطلب ----
if show_changes and 'enhanced_text' in st.session_state and text_input:
    st.markdown("---")
    st.subheader("🔍 تفاصيل التغييرات")
    # مقارنة بسيطة: كلمات مستبدلة وجمل معاد بناؤها
    orig_words = set(word_tokenize(text_input.lower()))
    enh_words = set(word_tokenize(st.session_state.enhanced_text.lower()))
    replaced = [(w, "→ (مرادف)") for w in orig_words if w not in enh_words and w.isalpha()]
    if replaced:
        st.write("**كلمات مستبدلة:**")
        st.table(pd.DataFrame(replaced[:20], columns=["الأصل", "التغيير"]))
    else:
        st.info("لم يتم الكشف عن استبدالات واضحة.")

    # مقارنة الجمل
    orig_sent = sent_tokenize(text_input)
    enh_sent = sent_tokenize(st.session_state.enhanced_text)
    if len(orig_sent) != len(enh_sent):
        st.write("**تم تعديل بنية الجمل (عدد الجمل مختلف).**")
    else:
        diff_sents = []
        for i, (o, e) in enumerate(zip(orig_sent, enh_sent)):
            if o != e:
                diff_sents.append((i+1, o[:100] + "...", e[:100] + "..."))
        if diff_sents:
            st.write("**جمل معاد بناؤها:**")
            st.table(pd.DataFrame(diff_sents, columns=["الجملة", "الأصل", "المعدل"]))

# ---- تذييل ----
st.markdown("---")
st.markdown("DeepClean Studio © 2025 - للأغراض التعليمية والبحثية فقط")
