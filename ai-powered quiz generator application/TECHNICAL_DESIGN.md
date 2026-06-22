# Technical Design — AI Quiz Generator

**Version:** 1.0  
**Date:** June 2026  
**Author:** AFIA-REFAL  
**Project:** TechVest Agentic AI — Academic Year 2025–26

---

## 1. Architecture Overview

AI Quiz Generator is a **single-file, single-process Streamlit application**. There is no separate backend server, database, or API layer. All logic — UI rendering, file parsing, AI calls, and state management — lives in `app.py` and executes within one Python process.

```
┌─────────────────────────────────────────────────────────────┐
│                        Browser (Client)                      │
│        Streamlit WebSocket + Static Assets                   │
└───────────────────────────────┬─────────────────────────────┘
                                │ HTTP / WS
┌───────────────────────────────▼─────────────────────────────┐
│                    Streamlit Server (Python)                  │
│  ┌──────────────┐  ┌────────────────┐  ┌──────────────────┐ │
│  │  UI Layer    │  │  Session State │  │  Backend Logic   │ │
│  │  (CSS + HTML)│  │  (st.session_  │  │  extract_pptx()  │ │
│  │  + Widgets   │  │   state dict)  │  │  llm()           │ │
│  └──────────────┘  └────────────────┘  │  gen_mcqs()      │ │
│                                        │  gen_exp()       │ │
│                                        └──────┬───────────┘ │
└─────────────────────────────────────────────┬─┘─────────────┘
                                              │ HTTPS POST
┌─────────────────────────────────────────────▼─────────────┐
│                  OpenRouter API                             │
│          https://openrouter.ai/api/v1/chat/completions     │
│          Model: anthropic/claude-haiku-4-5                 │
└────────────────────────────────────────────────────────────┘
```

---

## 2. Technology Choices

### 2.1 Streamlit

Streamlit was chosen because it allows a Python developer to build a fully interactive web application without writing any JavaScript. Every Python variable change triggers a full page re-render, making it natural to express a state-machine UI (upload → config → quiz → results) using `st.session_state`.

**Trade-offs:**
- Pro: Rapid development, no frontend build tooling required.
- Con: Full page re-render on every interaction. Worked around using `st.rerun()` only when state transitions are needed, not on every click.
- Con: Streamlit's default widget styling is opinionated. Overridden entirely with injected CSS targeting internal `data-testid` selectors.

### 2.2 OpenRouter / Claude Haiku

OpenRouter provides a single API endpoint that proxies multiple AI providers. `anthropic/claude-haiku-4-5` was selected for its balance of speed and instruction-following quality. MCQ generation requires the model to return structured JSON, which Haiku handles reliably with explicit prompt instructions.

**Trade-offs:**
- Pro: Single API key, model can be swapped by changing one constant.
- Con: Requires internet; no offline mode.
- Con: Response quality depends on prompt design (see Section 4).

### 2.3 python-pptx

The `python-pptx` library extracts text from every `TextFrame` across all shapes on all slides. It is pure Python with no native dependencies, making installation straightforward on all platforms including Python 3.14 using `--only-binary=:all:`.

---

## 3. Application State Machine

The application is driven by `st.session_state.step`, which acts as the current screen indicator. Navigation between screens is triggered by button clicks that mutate this value followed by `st.rerun()`.

```
             ┌──────────┐
   start ──► │  upload  │
             └────┬─────┘
      file parsed │
             ┌────▼─────┐
             │  config  │
             └────┬─────┘
    questions gen │
             ┌────▼─────┐
             │   quiz   │◄──── retake
             └────┬─────┘
        submitted │
             ┌────▼─────┐
             │ results  │──── upload new ──► upload (reset)
             └──────────┘
```

A second state variable, `st.session_state.nav_page`, controls which sidebar page is visible (`home` / `analytics` / `history` / `settings`) independently of the quiz flow.

### Full Session State Schema

| Key | Type | Description |
|---|---|---|
| `step` | `str` | Current quiz screen: `upload`, `config`, `quiz`, `results` |
| `nav_page` | `str` | Active sidebar tab: `home`, `analytics`, `history`, `settings` |
| `slides` | `list[dict]` | Parsed slide data: `[{slide: int, text: str}]` |
| `filename` | `str` | Uploaded file name |
| `slide_count` | `int` | Number of slides with extractable text |
| `word_count` | `int` | Total word count across all slides |
| `num_questions` | `int` | Configured question count (5–30) |
| `difficulty` | `str` | `Simple`, `Medium`, or `Complex` |
| `questions` | `list[dict]` | Generated MCQ objects from AI |
| `user_answers` | `dict[str, str]` | Maps question ID → selected option letter |
| `current_q` | `int` | Zero-based index of the question currently displayed |
| `start_time` | `float` | Unix timestamp when quiz started |
| `elapsed` | `int` | Total seconds taken (set on submit) |
| `explanations` | `dict[str, str]` | Maps question ID → AI explanation string |
| `rk` | `int` | Render key — incremented on retake to force button re-render |
| `quiz_history` | `list[dict]` | Records of completed quiz sessions |

---

## 4. Backend Functions

### 4.1 `extract_pptx(b: bytes) → tuple[list[dict], int]`

Parses a `.pptx` binary using `python-pptx`. Iterates over every slide, collects all paragraph text from all shapes that have a `TextFrame`, strips whitespace, joins into a single string per slide, and skips slides that produce no text (e.g. image-only slides).

Returns a list of `{slide: int, text: str}` dicts and the total word count.

```python
def extract_pptx(b):
    prs = Presentation(BytesIO(b))
    slides, words = [], 0
    for i, slide in enumerate(prs.slides, 1):
        txt = [p.text.strip()
               for s in slide.shapes if s.has_text_frame
               for p in s.text_frame.paragraphs if p.text.strip()]
        c = " ".join(txt)
        if c:
            slides.append({"slide": i, "text": c})
            words += len(c.split())
    return slides, words
```

### 4.2 `llm(prompt: str) → str`

Sends a single-turn chat completion request to the OpenRouter API. Uses `requests.post` with a 120-second timeout. Returns the raw text content of the first choice message.

Headers sent:
- `Authorization: Bearer <key>`
- `HTTP-Referer: http://localhost:8501` (required by OpenRouter)
- `X-Title: AI Quiz Generator`

### 4.3 `gen_mcqs(slides, n, diff) → list[dict]`

Constructs a prompt that embeds all slide text, specifies the difficulty via a pre-written instruction string, and instructs the model to return a JSON array. After receiving the response, strips any markdown code fences and parses with `json.loads`.

**MCQ schema returned by AI:**
```json
[
  {
    "id": 1,
    "question": "...",
    "options": {"A": "...", "B": "...", "C": "...", "D": "..."},
    "correct": "B",
    "topic": "label"
  }
]
```

**Difficulty prompts:**

| Level | Instruction |
|---|---|
| Simple | Factual recall — definitions, key terms, basic concepts directly from slides |
| Medium | Mix of factual recall and scenario-based reasoning |
| Complex | Analytical questions requiring deep understanding, edge-cases, trade-offs |

### 4.4 `gen_exp(questions, answers) → dict[str, str]`

Filters the question list to only wrong answers, constructs a block showing what the student selected vs. the correct answer, and asks the AI to produce a 1–2 sentence explanation per wrong answer. Returns a dict mapping question ID strings to explanation strings.

Returns `{}` if all answers were correct (no AI call needed).

---

## 5. UI Architecture

### 5.1 Layout Structure

The layout uses two Streamlit columns. The left column (6% width) renders the nav dock. The right column (94%) renders the header and content area.

```
st.columns([0.06, 0.94])
    ├── nav_col   → nav dock with 4 icon buttons
    └── main_col
        ├── top header HTML (title + step progress)
        └── st.columns([1, 3, 1])   ← centers the glass panel
               └── center col
                   └── screen content (upload / config / quiz / results)
                       or nav page content (analytics / history / settings)
```

### 5.2 CSS Strategy

All custom styling is injected via `st.markdown("<style>...</style>", unsafe_allow_html=True)` at app startup. Streamlit's internal elements are targeted using their `data-testid` attributes:

| Target | Selector |
|---|---|
| Main container | `[data-testid="block-container"]` |
| Streamlit header bar | `[data-testid="stHeader"]` |
| Default sidebar | `[data-testid="stSidebar"]` |
| Column wrappers | `[data-testid="column"]` |
| Vertical block wrappers | `[data-testid="stVerticalBlock"]` |
| File uploader | `[data-testid="stFileUploaderDropzone"]` |
| Slider | `[data-testid="stSlider"]` |
| Buttons | `.stButton > button[kind="primary|secondary"]` |

**Key CSS techniques used:**
- **Glassmorphism:** `backdrop-filter: blur(32px)` + `rgba` background on `.glass-panel`
- **Gradient mesh background:** Two `radial-gradient` layers on `stAppViewContainer`
- **Animated shimmer:** `@keyframes shimmer` with `background-position` sweep on score text
- **Floating icon:** `@keyframes floatIcon` using `translateY` on the upload icon
- **Pulse rings:** `@keyframes pulseRing` using `scale` + `opacity` for AI processing animation
- **Overflow fix:** All Streamlit column containers forced to `overflow: visible` to prevent card clipping

### 5.3 Option Selection Pattern

Streamlit does not natively support custom-styled radio button feedback. The solution uses `st.button` with `type="primary"` for the selected option and `type="secondary"` for all others, keyed with a `rk` (render key) that increments on retake to force React to unmount and re-mount buttons cleanly.

```python
for key, val in q["options"].items():
    picked = sel == key
    if st.button(f"{key}   {val}",
                 key=f"o_{i}_{key}_{st.session_state.rk}",
                 type="primary" if picked else "secondary"):
        st.session_state.user_answers[qid] = key
        st.rerun()
```

---

## 6. Data Flow

```
User uploads .pptx
        │
        ▼
extract_pptx(bytes)
  → list of {slide, text} dicts
        │
        ▼
User configures n, difficulty
        │
        ▼
gen_mcqs(slides, n, difficulty)
  → prompt construction
  → llm(prompt)
    → POST /chat/completions
    → raw JSON string response
  → json.loads(cleaned_response)
  → list of MCQ dicts stored in session_state.questions
        │
        ▼
User takes quiz
  → user_answers[qid] = selected_key per question
        │
        ▼
On submit:
  gen_exp(questions, user_answers)
  → filters wrong answers only
  → llm(explanation prompt)
  → dict of {qid: explanation}
  → stored in session_state.explanations
        │
        ▼
Results rendered from:
  session_state.questions
  session_state.user_answers
  session_state.explanations
```

---

## 7. API Integration

### 7.1 Request Format

```
POST https://openrouter.ai/api/v1/chat/completions
Content-Type: application/json
Authorization: Bearer <OPENROUTER_API_KEY>

{
  "model": "anthropic/claude-haiku-4-5",
  "messages": [{"role": "user", "content": "<prompt>"}],
  "temperature": 0.7,
  "max_tokens": 4096
}
```

### 7.2 Response Parsing

The model is instructed to return only raw JSON with no markdown fences. As a safety measure, the parser strips any ` ```json ` or ` ``` ` wrappers before calling `json.loads`. A `json.JSONDecodeError` is caught and surfaced to the user as a retry prompt.

### 7.3 Timeout

All requests use a 120-second timeout. For large slide decks (30+ slides), generation can take 20–40 seconds due to prompt size.

---

## 8. Error Handling

| Scenario | Handling |
|---|---|
| File > 25 MB | Checked before parsing; `st.error` shown |
| No text in slides | `extract_pptx` returns empty list; `st.error` shown |
| AI returns malformed JSON | `json.JSONDecodeError` caught; retry prompt shown |
| API HTTP error | `requests.HTTPError` caught; error message shown |
| API timeout | `requests.Timeout` caught via general `Exception`; error shown |
| Explanation generation fails | Silently returns `{}`; quiz results still display without explanations |

---

## 9. Security Considerations

- **API key storage:** Loaded from `.env` via `python-dotenv`; excluded from git via `.gitignore`.
- **No persistent storage:** All data lives in `st.session_state` (server-side Python dict, cleared on session end).
- **No user authentication:** Acceptable for local single-user deployment; would require auth layer for multi-user deployment.
- **Input validation:** File type and size validated before any processing occurs.
- **Prompt injection:** Slide text is embedded in prompts. No mitigation currently in place; acceptable given the controlled academic context.

---

## 10. Known Limitations

| Limitation | Impact | Mitigation |
|---|---|---|
| Image-only slides produce no text | Questions cannot be generated from visual-only content | Inform user via word count display |
| Session history lost on server restart | Analytics/history cleared | Acceptable for single-session use |
| Full page rerender on every interaction | Minor visual flicker on button clicks | Render key pattern minimises unnecessary rerenders |
| AI JSON output not schema-validated | Malformed responses surface as errors | Error message + retry; could add Pydantic validation |
| No question deduplication | Rare repeat questions on small decks | Acceptable; prompt instructs unique questions |

---

## 11. Deployment Notes

The application is designed for **local development use**. To deploy on a server:

1. Set `OPENROUTER_API_KEY` as an environment variable (not via `.env`).
2. Run with `streamlit run app.py --server.port 8501 --server.address 0.0.0.0`.
3. Consider adding Streamlit's built-in authentication or placing behind a reverse proxy with auth for multi-user environments.
4. Increase `--server.maxUploadSize` if files larger than 25 MB are needed.
