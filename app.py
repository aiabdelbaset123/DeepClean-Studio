#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepClean Studio - Claude API Edition
Uses Claude 3.5 Sonnet with aggressive humanization prompt.
Requires ANTHROPIC_API_KEY environment variable.
"""

from __future__ import annotations

import html
import json
import os
import re
import urllib.request
from dataclasses import dataclass
from io import BytesIO
from typing import List, Optional

import docx2txt
import pypdf
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(page_title="DeepClean Studio - Claude Humanizer", layout="wide")
AUTHOR_NAME = "Prof. Dr. Abdel-baset H. Mekky"

# ----------------------------------------------------------------------
# Claude System Prompt - The key to passing all detectors
# ----------------------------------------------------------------------
CLAUDE_SYSTEM = """You are a human academic editor with 20 years of experience. Rewrite the text to sound completely human.

CRITICAL RULES:

1. Every sentence must be complete. No fragments. No missing verbs.
2. Break long sentences (over 20 words) into two or three shorter ones.
3. NEVER use these words: additionally, moreover, furthermore, consequently, hence, crucial, pivotal, vital, significant, profound, robust, comprehensive, delve, showcase, underscore, highlight, resonate, align with, garner, tapestry, testament, landscape (abstract), intricate, multifaceted, serves as, stands as, marks a turning point, in conclusion, in summary, overall, it is important to note, not only, but also, uniquely, constitute, trajectories, pronounced, routinely, impose, reducing, exceeding, cumulative, dominant.
4. Replace them with simple words: also, so, but, and, important, key, big, helps, shows, is, part of, affects, cuts, above, often, brings.
5. Use short sentences (5-12 words) mixed with medium ones (12-20 words). Occasionally use a longer one (20-25 words) but break it with commas.
6. Start 15% of sentences with: So, Well, Look, Basically, I mean, You see.
7. Add "I think", "maybe", "probably", "it seems" before strong claims.
8. Keep all citations [1], [2] exactly as they are.
9. Keep all numbers (58.7 GW, 45°C, 2,400 kWh) exactly as they are.
10. Use active voice: "The IEA says" not "It is forecasted by the IEA".
11. Change passive constructions to active.
12. No markdown. No bold. No italics. No em dashes.
13. Output plain text only, with double line breaks between paragraphs.

Return ONLY the rewritten text. No explanations. No summaries. No commentary."""

def call_claude_api(text: str, intensity: int) -> str:
    """Call Claude API with the humanization prompt."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not set in environment variables")
    
    intensity_note = {
        1: "Minimal changes. Just remove forbidden words and fix grammar.",
        2: "Light rewrite. Shorten long sentences, remove forbidden words.",
        3: "Moderate rewrite. Shorten long sentences, add occasional human touches.",
        4: "Substantial rewrite. Aggressively shorten sentences, add 'I think', 'so', 'well'.",
        5: "Maximum rewrite. Make it sound like a busy researcher writing quickly."
    }.get(intensity, "Moderate rewrite.")
    
    user_message = f"Intensity: {intensity_note}\n\nText:\n{text}"
    
    payload = json.dumps({
        "model": "claude-3-5-sonnet-20241022",
        "max_tokens": 2000,
        "temperature": 0.7 + (intensity - 1) * 0.07,  # more randomness at higher intensity
        "system": CLAUDE_SYSTEM,
        "messages": [{"role": "user", "content": user_message}]
    }).encode("utf-8")
    
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        },
        method="POST"
    )
    
    with urllib.request.urlopen(req, timeout=90) as response:
        data = json.loads(response.read().decode("utf-8"))
        text_blocks = [block["text"] for block in data.get("content", []) if block.get("type") == "text"]
        return "\n".join(text_blocks).strip()


# ----------------------------------------------------------------------
# Fallback local engine (simple rule-based, used if Claude fails)
# ----------------------------------------------------------------------
class SimpleFallbackEngine:
    def __init__(self, text: str):
        self.text = text
    
    def run(self) -> str:
        # Just remove obvious AI markers
        text = self.text
        text = re.sub(r'(?i)\b(additionally|moreover|furthermore|consequently|hence|crucial|pivotal|vital|significant|profound|robust|comprehensive|delve|showcase|underscore|highlight|resonate|align with|garner|tapestry|testament|landscape|intricate|multifaceted|serves as|stands as|constitute|trajectories|pronounced|routinely|impose|reducing|exceeding|cumulative|dominant)\b', '', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()


# ----------------------------------------------------------------------
# UI Helpers
# ----------------------------------------------------------------------
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)

def extract_uploaded(uploaded) -> str:
    name = uploaded.name.lower()
    if name.endswith(".txt"):
        raw = uploaded.read()
        try:
            return raw.decode("utf-8")
        except:
            return raw.decode("utf-8-sig")
    if name.endswith(".docx"):
        return docx2txt.process(uploaded) or ""
    if name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""

def make_word_file(text: str, title: Optional[str] = None) -> BytesIO:
    doc = Document()
    doc.core_properties.author = AUTHOR_NAME
    doc.core_properties.title = title or "DeepClean Humanized"
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
        run.font.size = Pt(14)
        run.bold = True
    
    for line in text.split("\n"):
        if line.strip():
            p = doc.add_paragraph(line.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.25)
        else:
            doc.add_paragraph()
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ----------------------------------------------------------------------
# Main App
# ----------------------------------------------------------------------
def main():
    st.title("🤖 DeepClean Studio – Claude Humanizer")
    st.caption("يستخدم Claude API لإعادة كتابة النصوص بأسلوب بشري حقيقي – يجتاز GPTZero، ZeroGPT، Originality.ai")
    st.caption(AUTHOR_NAME)
    
    # API key input
    api_key = st.sidebar.text_input("🔑 ANTHROPIC_API_KEY", type="password", 
                                     help="احصل على مفتاح من console.anthropic.com")
    if api_key:
        os.environ["ANTHROPIC_API_KEY"] = api_key
    
    with st.sidebar:
        st.header("⚙️ الإعدادات")
        source = st.radio("مصدر النص", ("📄 لصق نص", "📁 رفع ملف"))
        user_text = ""
        
        if source == "📁 رفع ملف":
            uploaded = st.file_uploader("اختر ملفًا", type=["txt", "docx", "pdf"])
            if uploaded:
                user_text = extract_uploaded(uploaded)
        else:
            user_text = st.text_area("ألصق النص الأكاديمي هنا", height=250)
        
        intensity = st.slider("قوة المراجعة", 1, 5, 3, 
                              help="1=تغييرات بسيطة، 5=مراجعة جذرية بأسلوب بشري")
        
        use_api = st.checkbox("✅ استخدام Claude API (موصى بشدة)", value=True)
        
        if st.button("🚀 بدء المراجعة", type="primary", use_container_width=True):
            if not user_text:
                st.warning("الرجاء إدخال نص أو رفع ملف.")
            elif use_api and not api_key:
                st.error("يرجى إدخال مفتاح API الخاص بـ Claude.")
            else:
                with st.spinner("جاري إعادة الكتابة بأسلوب بشري..."):
                    try:
                        if use_api and api_key:
                            revised = call_claude_api(user_text, intensity)
                        else:
                            st.info("استخدام المحرك المحلي (أقل فعالية). يوصى بـ Claude API.")
                            engine = SimpleFallbackEngine(user_text)
                            revised = engine.run()
                        st.session_state.revised = revised
                        st.session_state.original = user_text
                        st.session_state.done = True
                    except Exception as e:
                        st.error(f"خطأ: {e}")
    
    colA, colB = st.columns(2)
    
    with colA:
        st.subheader("📄 النص الأصلي")
        if user_text:
            st.text_area("", user_text, height=450, key="orig_area")
            st.caption(f"كلمات: {len(tokenize_words(user_text))}")
        else:
            st.info("أدخل نصًا من الشريط الجانبي.")
    
    with colB:
        st.subheader("✨ النص المعاد كتابته (بشري)")
        if st.session_state.get("done") and st.session_state.get("revised"):
            rev = st.session_state.revised
            st.markdown(f"<div style='background:#f5f5f0; padding:15px; border-radius:8px; font-family:Times New Roman; font-size:12pt;'>{html.escape(rev).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
            st.text_area("", rev, height=450, key="rev_area", label_visibility="collapsed")
            st.caption(f"كلمات: {len(tokenize_words(rev))}")
            word_file = make_word_file(rev, "DeepClean_Humanized")
            st.download_button("📥 تحميل Word", data=word_file, file_name="humanized.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        else:
            st.info("ستظهر النسخة البشرية هنا بعد المعالجة.")

if __name__ == "__main__":
    if "done" not in st.session_state:
        st.session_state.done = False
    if "revised" not in st.session_state:
        st.session_state.revised = ""
    if "original" not in st.session_state:
        st.session_state.original = ""
    main()
