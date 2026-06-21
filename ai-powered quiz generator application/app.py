import os
import json
import re
import time
import requests
import streamlit as st
from pptx import Presentation
from dotenv import load_dotenv
from io import BytesIO

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_URL    = "https://openrouter.ai/api/v1/chat/completions"
MODEL             = "anthropic/claude-haiku-4-5"

DIFF_HINTS = {
    "Simple":  "Definitions, key terms & basic concepts — straight from the slides.",
    "Medium":  "Mix of recall + scenario reasoning spread across all topics.",
    "Complex": "Analytical edge-cases, trade-offs & 'best-among-goods' challenges.",
}
DIFF_PROMPTS = {
    "Simple":  "Generate straightforward factual recall questions. Focus on definitions, key terms, and basic concepts directly stated in the slides.",
    "Medium":  "Generate balanced questions mixing factual recall and scenario-based reasoning. Include some 'best answer' questions that require understanding relationships between concepts.",
    "Complex": "Generate challenging analytical questions requiring deep understanding. Focus on edge cases, subtle distinctions, trade-offs, and application of concepts in novel scenarios. Make distractors highly plausible.",
}

st.set_page_config(page_title="AI Quiz Generator", page_icon="🎯", layout="centered")

# ── COMPLETE UI REDESIGN ────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', -apple-system, sans-serif;
    background: linear-gradient(135deg, #0f0c29 0%, #1a1040 40%, #0d1b4b 100%);
    min-height: 100vh;
    color: #fff;
}

[data-testid="stHeader"] { background: transparent !important; }
#MainMenu, footer, header { visibility: hidden; }
section[data-testid="stSidebar"] { display: none; }

[data-testid="stMain"] {
    padding-top: 0 !important;
}

/* ── NAVBAR ── */
.navbar {
    background: rgba(255,255,255,0.04);
    backdrop-filter: blur(20px);
    border-bottom: 1px solid rgba(255,255,255,0.08);
    padding: 14px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 32px;
    border-radius: 0 0 16px 16px;
}
.navbar-brand {
    display: flex; align-items: center; gap: 10px;
    font-size: 18px; font-weight: 800; color: #fff;
    letter-spacing: -0.3px;
}
.navbar-logo {
    width: 36px; height: 36px; border-radius: 10px;
    background: linear-gradient(135deg, #a78bfa, #6366f1);
    display: flex; align-items: center; justify-content: center;
    font-size: 18px; font-weight: 900; color: #fff;
    box-shadow: 0 4px 15px rgba(99,102,241,0.5);
}
.navbar-step {
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.15);
    color: rgba(255,255,255,0.7);
    font-size: 12px; font-weight: 600;
    padding: 5px 14px; border-radius: 20px;
    letter-spacing: 0.05em;
}

/* ── GLASS CARD ── */
.glass-card {
    background: rgba(255,255,255,0.05);
    backdrop-filter: blur(24px);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 20px;
    padding: 36px;
    margin-bottom: 20px;
    box-shadow: 0 8px 40px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.1);
}

/* ── UPLOAD SCREEN ── */
.upload-zone {
    border: 2px dashed rgba(167,139,250,0.5);
    border-radius: 16px;
    padding: 52px 28px;
    text-align: center;
    background: rgba(99,102,241,0.05);
    margin-bottom: 16px;
    transition: all .2s;
}
.upload-zone:hover { border-color: #a78bfa; background: rgba(99,102,241,0.1); }
.upload-icon { font-size: 52px; margin-bottom: 14px; display: block; }
.upload-zone h2 {
    font-size: 20px; font-weight: 700; color: #fff;
    margin-bottom: 6px;
}
.upload-zone p { font-size: 13px; color: rgba(255,255,255,0.45); }

.file-parsed {
    background: rgba(16,185,129,0.1);
    border: 1px solid rgba(16,185,129,0.3);
    border-radius: 14px;
    padding: 16px 18px;
    display: flex; align-items: center; gap: 14px;
    margin: 14px 0;
}
.file-parsed .parsed-icon { font-size: 24px; }
.file-parsed .parsed-info { flex: 1; }
.file-parsed .parsed-info strong { color: #fff; font-size: 14px; display: block; margin-bottom: 2px; }
.file-parsed .parsed-info span { color: rgba(255,255,255,0.5); font-size: 12px; }
.parsed-badge {
    background: linear-gradient(135deg, #10b981, #059669);
    color: #fff; font-size: 11px; font-weight: 700;
    padding: 4px 12px; border-radius: 20px;
    letter-spacing: .03em;
}

/* ── CONFIGURE SCREEN ── */
.source-pill {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(167,139,250,0.15);
    border: 1px solid rgba(167,139,250,0.3);
    border-radius: 10px;
    padding: 10px 16px;
    font-size: 13px; color: #c4b5fd;
    font-weight: 500;
    margin-bottom: 24px;
    width: 100%;
}

.section-label {
    font-size: 11px; font-weight: 700; letter-spacing: .1em;
    color: rgba(255,255,255,0.4);
    text-transform: uppercase;
    margin-bottom: 10px; display: block;
}

.diff-grid { display: flex; gap: 10px; margin-bottom: 10px; }
.diff-card {
    flex: 1; padding: 14px 10px; border-radius: 12px;
    border: 1.5px solid rgba(255,255,255,0.1);
    background: rgba(255,255,255,0.04);
    text-align: center; cursor: pointer;
    transition: all .2s;
}
.diff-card:hover { border-color: #a78bfa; background: rgba(167,139,250,0.1); }
.diff-card.active {
    border-color: #a78bfa;
    background: linear-gradient(135deg, rgba(167,139,250,0.2), rgba(99,102,241,0.2));
    box-shadow: 0 0 20px rgba(167,139,250,0.2);
}
.diff-card .diff-emoji { font-size: 22px; margin-bottom: 6px; display: block; }
.diff-card .diff-name { font-size: 13px; font-weight: 700; color: #fff; }

.diff-hint-box {
    background: rgba(167,139,250,0.08);
    border: 1px solid rgba(167,139,250,0.2);
    border-radius: 10px; padding: 10px 14px;
    font-size: 12px; color: #c4b5fd;
    margin-bottom: 20px;
}

/* ── QUIZ SCREEN ── */
.quiz-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 8px;
}
.q-counter {
    font-size: 12px; font-weight: 700; color: #a78bfa;
    letter-spacing: .08em;
}
.timer-chip {
    background: rgba(251,146,60,0.15);
    border: 1px solid rgba(251,146,60,0.3);
    color: #fb923c; font-size: 13px; font-weight: 700;
    padding: 5px 12px; border-radius: 20px;
    display: inline-flex; align-items: center; gap: 5px;
}

.progress-track {
    background: rgba(255,255,255,0.08);
    border-radius: 4px; height: 5px;
    margin-bottom: 24px; overflow: hidden;
}
.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #a78bfa, #6366f1);
    border-radius: 4px;
    transition: width .4s ease;
}

.question-text {
    font-size: 18px; font-weight: 700; color: #fff;
    line-height: 1.5; margin-bottom: 22px;
}

.opt-wrap { display: flex; flex-direction: column; gap: 10px; margin-bottom: 24px; }
.opt-item {
    display: flex; align-items: center; gap: 14px;
    background: rgba(255,255,255,0.04);
    border: 1.5px solid rgba(255,255,255,0.09);
    border-radius: 12px; padding: 14px 16px;
    cursor: pointer; transition: all .15s;
    font-size: 14px; color: rgba(255,255,255,0.85);
}
.opt-item:hover { border-color: #a78bfa; background: rgba(167,139,250,0.1); }
.opt-item.picked {
    border-color: #a78bfa;
    background: linear-gradient(135deg, rgba(167,139,250,0.18), rgba(99,102,241,0.12));
    color: #fff;
    box-shadow: 0 0 0 3px rgba(167,139,250,0.15);
}
.opt-badge {
    width: 30px; height: 30px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px; font-weight: 800; flex-shrink: 0;
    background: rgba(255,255,255,0.08);
    border: 1.5px solid rgba(255,255,255,0.12);
    color: rgba(255,255,255,0.6);
}
.opt-item.picked .opt-badge {
    background: linear-gradient(135deg, #a78bfa, #6366f1);
    border-color: transparent; color: #fff;
}

/* ── RESULTS SCREEN ── */
.score-hero {
    background: linear-gradient(135deg, rgba(167,139,250,0.15), rgba(99,102,241,0.1));
    border: 1px solid rgba(167,139,250,0.25);
    border-radius: 18px; padding: 28px;
    text-align: center; margin-bottom: 24px;
}
.score-number {
    font-size: 64px; font-weight: 900; line-height: 1;
    background: linear-gradient(135deg, #a78bfa, #6366f1);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin-bottom: 4px;
}
.score-pct {
    font-size: 22px; font-weight: 700; color: #fff;
    margin-bottom: 10px;
}
.score-headline { font-size: 16px; color: rgba(255,255,255,0.65); margin-bottom: 16px; }
.stat-row { display: flex; justify-content: center; gap: 12px; flex-wrap: wrap; }
.stat-chip {
    padding: 6px 16px; border-radius: 20px;
    font-size: 13px; font-weight: 600;
}
.stat-correct { background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); color: #34d399; }
.stat-wrong   { background: rgba(239,68,68,0.12);  border: 1px solid rgba(239,68,68,0.25);  color: #f87171; }
.stat-time    { background: rgba(251,146,60,0.12);  border: 1px solid rgba(251,146,60,0.25);  color: #fb923c; }

.review-section-label {
    font-size: 11px; font-weight: 700; letter-spacing: .1em;
    color: rgba(255,255,255,0.35); text-transform: uppercase;
    margin: 20px 0 12px;
}
.rev-card {
    border-radius: 14px; padding: 15px 16px;
    margin-bottom: 10px;
}
.rev-card.ok {
    background: rgba(16,185,129,0.08);
    border: 1px solid rgba(16,185,129,0.22);
}
.rev-card.err {
    background: rgba(239,68,68,0.07);
    border: 1px solid rgba(239,68,68,0.2);
}
.rev-header {
    display: flex; align-items: flex-start; gap: 10px; margin-bottom: 6px;
}
.rev-icon { font-size: 16px; flex-shrink: 0; margin-top: 2px; }
.rev-title { font-size: 13px; font-weight: 700; color: #fff; }
.rev-q     { font-size: 12px; color: rgba(255,255,255,0.55); margin: 3px 0 0 26px; }
.rev-ans   { font-size: 12px; margin: 4px 0 0 26px; }
.rev-ans.ok-ans  { color: #34d399; }
.rev-ans.err-ans { color: #f87171; }
.ai-box {
    background: rgba(167,139,250,0.08);
    border: 1px solid rgba(167,139,250,0.2);
    border-radius: 10px; padding: 10px 12px;
    font-size: 12px; color: #c4b5fd;
    margin: 8px 0 0 0;
    display: flex; gap: 8px; align-items: flex-start;
}

/* ── STREAMLIT OVERRIDES ── */
div[data-testid="stFileUploaderDropzone"] {
    background: rgba(99,102,241,0.06) !important;
    border: 2px dashed rgba(167,139,250,0.4) !important;
    border-radius: 14px !important;
    color: rgba(255,255,255,0.6) !important;
}
div[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #a78bfa !important;
    background: rgba(99,102,241,0.12) !important;
}
div[data-testid="stFileUploaderDropzone"] label { color: #c4b5fd !important; }

/* Slider */
div[data-testid="stSlider"] > div > div > div > div {
    background: linear-gradient(90deg, #a78bfa, #6366f1) !important;
}
div[data-testid="stSlider"] label { color: rgba(255,255,255,0.8) !important; font-weight: 600 !important; }
div[data-testid="stSlider"] p { color: rgba(255,255,255,0.5) !important; }

/* All buttons */
.stButton > button {
    width: 100% !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    padding: 13px 20px !important;
    transition: all .2s !important;
    letter-spacing: .01em !important;
    text-align: left !important;
    justify-content: flex-start !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #a78bfa, #6366f1) !important;
    color: #fff !important;
    border: none !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.4) !important;
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 6px 28px rgba(99,102,241,0.6) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="secondary"] {
    background: rgba(255,255,255,0.05) !important;
    border: 1.5px solid rgba(255,255,255,0.12) !important;
    color: rgba(255,255,255,0.8) !important;
}
.stButton > button[kind="secondary"]:hover {
    background: rgba(167,139,250,0.1) !important;
    border-color: #a78bfa !important;
    color: #fff !important;
}
.stButton > button:disabled {
    opacity: 0.35 !important;
    cursor: not-allowed !important;
    transform: none !important;
    box-shadow: none !important;
}

/* Spinner */
div[data-testid="stSpinner"] > div { border-top-color: #a78bfa !important; }

/* Error / success */
div[data-testid="stAlert"] {
    border-radius: 12px !important;
    border: none !important;
}

/* Captions */
.stCaption, [data-testid="stCaptionContainer"] p {
    color: rgba(255,255,255,0.4) !important;
    font-size: 12px !important;
}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "step": "upload",
        "slides": [], "filename": "", "slide_count": 0,
        "word_count": 0, "preview": "",
        "num_questions": 10, "difficulty": "Medium",
        "questions": [], "user_answers": {},
        "current_q": 0, "start_time": None,
        "elapsed": 0, "explanations": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── BACKEND ──────────────────────────────────────────────────────────────────
def extract_pptx(file_bytes):
    prs = Presentation(BytesIO(file_bytes))
    slides, total_words = [], 0
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
        combined = " ".join(texts)
        if combined:
            slides.append({"slide": i, "text": combined})
            total_words += len(combined.split())
    return slides, total_words

def call_llm(prompt):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "http://localhost:8501",
        "X-Title": "AI Quiz Generator",
    }
    r = requests.post(OPENROUTER_URL, headers=headers, json={
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7, "max_tokens": 4096,
    }, timeout=120)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]

def generate_mcqs(slides, num_q, difficulty):
    slide_text = "\n\n".join(f"[Slide {s['slide']}]: {s['text']}" for s in slides)
    prompt = f"""You are an expert quiz designer. Based on the slide content below, generate exactly {num_q} multiple-choice questions.

DIFFICULTY: {difficulty}
INSTRUCTION: {DIFF_PROMPTS[difficulty]}

SLIDE CONTENT:
{slide_text}

RULES:
- Exactly 4 options per question labeled A, B, C, D
- Only one correct answer; three plausible distractors
- Balance questions across all slides/topics
- Favor scenario-based questions over pure recall

Return ONLY a valid JSON array — no markdown fences, no extra text:
[{{"id":1,"question":"...","options":{{"A":"...","B":"...","C":"...","D":"..."}},"correct":"B","topic":"Brief label"}}]"""
    raw = call_llm(prompt)
    raw = re.sub(r"```(?:json)?", "", raw).strip().strip("`").strip()
    return json.loads(raw)

def generate_explanations(questions, user_answers):
    wrong = [q for q in questions
             if user_answers.get(str(q["id"])) and user_answers[str(q["id"])] != q["correct"]]
    if not wrong:
        return {}
    block = "\n\n".join(
        f"Q{w['id']}: {w['question']}\n"
        f"Student: {user_answers[str(w['id'])]} — {w['options'].get(user_answers[str(w['id'])],'')}\n"
        f"Correct: {w['correct']} — {w['options'][w['correct']]}"
        for w in wrong
    )
    raw = call_llm(f"""For each wrong answer write a concise 1-2 sentence explanation of why the student's answer is wrong and the correct answer is right.

{block}

Return ONLY JSON: {{"1":"explanation","5":"explanation"}}""")
    raw = re.sub(r"```(?:json)?", "", raw).strip().strip("`").strip()
    return json.loads(raw)

# ── NAVBAR ───────────────────────────────────────────────────────────────────
step_label = {
    "upload":  "Step 1 of 3  ·  Upload",
    "config":  "Step 2 of 3  ·  Configure",
    "quiz":    f"Q {st.session_state.current_q+1} / {len(st.session_state.questions)}",
    "results": "Quiz Complete ✓",
}.get(st.session_state.step, "")

st.markdown(f"""
<div class="navbar">
  <div class="navbar-brand">
    <div class="navbar-logo">🎯</div>
    AI Quiz Generator
  </div>
  <span class="navbar-step">{step_label}</span>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SCREEN 1 — UPLOAD
# ════════════════════════════════════════════════════════════════════════════
if st.session_state.step == "upload":

    st.markdown("""
    <div style="text-align:center;margin-bottom:28px">
      <div style="font-size:13px;font-weight:600;letter-spacing:.1em;color:rgba(255,255,255,.35);text-transform:uppercase;margin-bottom:8px">Step 1 of 3</div>
      <h1 style="font-size:28px;font-weight:900;color:#fff;letter-spacing:-.5px;margin-bottom:6px">Upload Your Presentation</h1>
      <p style="font-size:14px;color:rgba(255,255,255,.45)">Drop your .pptx file and we'll extract all the content automatically</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload your PowerPoint file", type=["pptx"],
        label_visibility="collapsed"
    )

    if uploaded:
        if not uploaded.name.lower().endswith(".pptx"):
            st.error("❌  Only .pptx files are supported.")
        elif uploaded.size > 25 * 1024 * 1024:
            st.error("❌  File exceeds the 25 MB limit.")
        else:
            with st.spinner("Parsing slides…"):
                try:
                    slides, word_count = extract_pptx(uploaded.read())
                    if not slides:
                        st.error("No text content found in this file.")
                    else:
                        st.session_state.slides      = slides
                        st.session_state.filename    = uploaded.name
                        st.session_state.slide_count = len(slides)
                        st.session_state.word_count  = word_count
                        st.session_state.preview     = slides[0]["text"][:220]
                        st.markdown(f"""
                        <div class="file-parsed">
                          <span class="parsed-icon">📄</span>
                          <div class="parsed-info">
                            <strong>{uploaded.name}</strong>
                            <span>{len(slides)} slides parsed · {word_count:,} words extracted</span>
                          </div>
                          <span class="parsed-badge">✓ Ready</span>
                        </div>
                        """, unsafe_allow_html=True)
                        st.caption(f"Preview — {st.session_state.preview}…")
                except Exception as e:
                    st.error(f"Failed to parse: {e}")

    st.write("")
    if st.button("Continue →", type="primary", disabled=not bool(st.session_state.slides)):
        st.session_state.step = "config"
        st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# SCREEN 2 — CONFIGURE
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.step == "config":

    st.markdown("""
    <div style="text-align:center;margin-bottom:28px">
      <div style="font-size:13px;font-weight:600;letter-spacing:.1em;color:rgba(255,255,255,.35);text-transform:uppercase;margin-bottom:8px">Step 2 of 3</div>
      <h1 style="font-size:28px;font-weight:900;color:#fff;letter-spacing:-.5px;margin-bottom:6px">Configure Your Quiz</h1>
      <p style="font-size:14px;color:rgba(255,255,255,.45)">Choose how many questions and how challenging you want them</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="source-pill">
      📎 &nbsp;<strong style="color:#fff">{st.session_state.filename}</strong>
      &nbsp;·&nbsp; {st.session_state.slide_count} slides &nbsp;·&nbsp; {st.session_state.word_count:,} words
    </div>
    """, unsafe_allow_html=True)

    st.session_state.num_questions = st.slider(
        "Number of questions", min_value=5, max_value=30,
        value=st.session_state.num_questions, step=1,
    )

    st.markdown('<span class="section-label">Difficulty Level</span>', unsafe_allow_html=True)

    DIFF_META = {
        "Simple":  ("🟢", "Simple"),
        "Medium":  ("🟡", "Medium"),
        "Complex": ("🔴", "Complex"),
    }
    col1, col2, col3 = st.columns(3)
    for col, (key, (emoji, name)) in zip([col1, col2, col3], DIFF_META.items()):
        with col:
            active = "active" if st.session_state.difficulty == key else ""
            st.markdown(f"""
            <div class="diff-card {active}" id="dc_{key}">
              <span class="diff-emoji">{emoji}</span>
              <span class="diff-name">{name}</span>
            </div>
            """, unsafe_allow_html=True)
            if st.button(name, key=f"d_{key}",
                         type="primary" if st.session_state.difficulty == key else "secondary"):
                st.session_state.difficulty = key
                st.rerun()

    st.markdown(f"""
    <div class="diff-hint-box">
      💡 &nbsp;{DIFF_HINTS[st.session_state.difficulty]}
    </div>
    """, unsafe_allow_html=True)

    if st.button(f"⚡  Generate {st.session_state.num_questions} Questions", type="primary"):
        with st.spinner("AI is crafting your quiz — this takes ~15 seconds…"):
            try:
                qs = generate_mcqs(
                    st.session_state.slides,
                    st.session_state.num_questions,
                    st.session_state.difficulty,
                )
                st.session_state.questions   = qs
                st.session_state.user_answers = {}
                st.session_state.current_q   = 0
                st.session_state.start_time  = time.time()
                st.session_state.step        = "quiz"
                st.rerun()
            except json.JSONDecodeError:
                st.error("AI returned a malformed response — please try again.")
            except Exception as e:
                st.error(f"Generation failed: {e}")

# ════════════════════════════════════════════════════════════════════════════
# SCREEN 3 — QUIZ
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.step == "quiz":
    questions = st.session_state.questions
    total     = len(questions)
    i         = st.session_state.current_q
    q         = questions[i]
    qid       = str(q["id"])

    elapsed      = int(time.time() - (st.session_state.start_time or time.time()))
    mins, secs   = divmod(elapsed, 60)
    pct          = int(((i + 1) / total) * 100)
    answered     = sum(1 for qq in questions if str(qq["id"]) in st.session_state.user_answers)

    # Header row
    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown(f'<div class="q-counter">QUESTION {i+1} OF {total} &nbsp;·&nbsp; {answered} answered</div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="timer-chip">⏱ {mins:02d}:{secs:02d}</div>', unsafe_allow_html=True)

    # Progress bar
    st.markdown(f"""
    <div class="progress-track">
      <div class="progress-fill" style="width:{pct}%"></div>
    </div>
    """, unsafe_allow_html=True)

    # Question
    st.markdown(f'<div class="question-text">{q["question"]}</div>', unsafe_allow_html=True)

    # Options — single Streamlit button per choice, styled via CSS
    selected = st.session_state.user_answers.get(qid)
    for key in ["A", "B", "C", "D"]:
        is_picked = selected == key
        if st.button(
            f"{key}   {q['options'][key]}",
            key=f"btn_{i}_{key}",
            use_container_width=True,
            type="primary" if is_picked else "secondary",
        ):
            st.session_state.user_answers[qid] = key
            st.rerun()

    st.write("")
    c_prev, c_next = st.columns(2)
    with c_prev:
        if i > 0:
            if st.button("← Previous", type="secondary"):
                st.session_state.current_q -= 1
                st.rerun()
    with c_next:
        is_last = (i == total - 1)
        lbl = "Submit Quiz ✓" if is_last else "Next →"
        if st.button(lbl, type="primary"):
            if not is_last:
                st.session_state.current_q += 1
                st.rerun()
            else:
                st.session_state.elapsed = int(time.time() - st.session_state.start_time)
                with st.spinner("Generating AI explanations for wrong answers…"):
                    try:
                        st.session_state.explanations = generate_explanations(
                            questions, st.session_state.user_answers)
                    except Exception:
                        st.session_state.explanations = {}
                st.session_state.step = "results"
                st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# SCREEN 4 — RESULTS
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.step == "results":
    questions = st.session_state.questions
    total     = len(questions)
    correct   = sum(1 for q in questions
                    if st.session_state.user_answers.get(str(q["id"])) == q["correct"])
    wrong     = total - correct
    pct       = round((correct / total) * 100) if total else 0
    mins, secs = divmod(st.session_state.elapsed, 60)

    headline = (
        "Perfect score! 🎉" if pct == 100 else
        "Outstanding work! 🌟" if pct >= 90 else
        "Great job! Keep it up 💪" if pct >= 80 else
        "Good effort — review the misses 📚" if pct >= 60 else
        "Keep studying — you'll get there 🚀"
    )

    # Score hero
    st.markdown(f"""
    <div style="text-align:center;margin-bottom:10px">
      <div style="font-size:13px;font-weight:600;letter-spacing:.1em;color:rgba(255,255,255,.35);text-transform:uppercase;margin-bottom:8px">Results</div>
      <h1 style="font-size:26px;font-weight:900;color:#fff;letter-spacing:-.5px">{headline}</h1>
    </div>
    <div class="score-hero">
      <div class="score-number">{correct}/{total}</div>
      <div class="score-pct">{pct}%</div>
      <div class="score-headline">{st.session_state.difficulty} difficulty · {st.session_state.slide_count}-slide deck</div>
      <div class="stat-row">
        <span class="stat-chip stat-correct">✓ {correct} correct</span>
        <span class="stat-chip stat-wrong">✗ {wrong} wrong</span>
        <span class="stat-chip stat-time">⏱ {mins:02d}:{secs:02d}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="review-section-label">Review &amp; AI Feedback</div>', unsafe_allow_html=True)

    for idx, q in enumerate(questions):
        qid      = str(q["id"])
        user_ans = st.session_state.user_answers.get(qid)
        ok       = user_ans == q["correct"]
        card_cls = "ok" if ok else "err"
        icon     = "✅" if ok else "❌"

        if ok:
            ans_html = f'<span class="rev-ans ok-ans">Your answer: {user_ans} ({q["options"].get(user_ans,"")}) — correct</span>'
        elif user_ans:
            ans_html = (f'<span class="rev-ans err-ans">'
                        f'You chose {user_ans}: {q["options"].get(user_ans,"")} &nbsp;·&nbsp; '
                        f'Correct: {q["correct"]}: {q["options"][q["correct"]]}</span>')
        else:
            ans_html = (f'<span class="rev-ans err-ans">'
                        f'Not answered &nbsp;·&nbsp; Correct: {q["correct"]}: {q["options"][q["correct"]]}</span>')

        exp = st.session_state.explanations.get(qid, "")
        ai_html = ""
        if not ok and exp:
            ai_html = f'<div class="ai-box"><span>🤖</span><span>{exp}</span></div>'

        st.markdown(f"""
        <div class="rev-card {card_cls}">
          <div class="rev-header">
            <span class="rev-icon">{icon}</span>
            <span class="rev-title">Q{idx+1} &nbsp;·&nbsp; {q.get("topic","")}</span>
          </div>
          <div class="rev-q">{q["question"]}</div>
          {ans_html}
          {ai_html}
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↺  Retake Quiz", type="primary"):
            st.session_state.user_answers  = {}
            st.session_state.current_q     = 0
            st.session_state.start_time    = time.time()
            st.session_state.explanations  = {}
            st.session_state.step          = "quiz"
            st.rerun()
    with c2:
        if st.button("↑  Upload New PPT", type="secondary"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
