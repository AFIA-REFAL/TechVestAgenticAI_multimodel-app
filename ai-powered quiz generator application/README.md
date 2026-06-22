# AI Quiz Generator

> Transform any PowerPoint presentation into an interactive AI-powered quiz in seconds.

---

## Overview

AI Quiz Generator is a Streamlit-based web application that accepts a `.pptx` file and uses a large language model to automatically generate multiple-choice questions (MCQs) from the slide content. Users configure the number of questions and difficulty level, take the quiz in a clean one-question-per-screen interface, and receive a detailed results breakdown with AI-generated explanations for every incorrect answer.

---

## Features

| Feature | Description |
|---|---|
| **Smart Question Generation** | Claude Haiku (via OpenRouter) reads slide text and generates contextually accurate MCQs |
| **3 Difficulty Levels** | Simple (recall), Medium (reasoning), Complex (analytical / edge-case) |
| **5–30 Questions** | Configurable slider — generate as few or as many as needed |
| **One Question Per Screen** | Distraction-free quiz experience with progress bar and live timer |
| **AI Explanations** | Wrong answers receive a 1–2 sentence AI explanation on the results screen |
| **Quiz History** | Every completed session is saved and viewable in the History tab |
| **Analytics Dashboard** | Session-level stats: avg score, best score, total questions answered |
| **Premium Dark UI** | Glassmorphism cards, indigo/violet gradients, animated upload zone |

---

## Screenshots

```
Upload → Configure → Quiz → Results
```

- **Upload screen** — animated drag-and-drop zone with AI processing animation
- **Configure screen** — slider + 3-tile difficulty selector
- **Quiz screen** — single question with A/B/C/D option cards, progress bar, timer
- **Results screen** — score banner, per-question review, AI feedback bubbles

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend / UI | Streamlit 1.35+ with custom CSS (glassmorphism dark theme) |
| AI Model | `anthropic/claude-haiku-4-5` via OpenRouter API |
| File Parsing | `python-pptx` |
| HTTP Client | `requests` |
| Config | `python-dotenv` |
| Language | Python 3.10+ |

---

## Getting Started

### Prerequisites

- Python 3.10 or higher
- An [OpenRouter](https://openrouter.ai) API key

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/AFIA-REFAL/TechVestAgenticAI.git
cd TechVestAgenticAI/ai-powered\ quiz\ generator\ application

# 2. Install dependencies
pip install -r requirements.txt --only-binary=:all:
```

### Configuration

Create a `.env` file in the `ai-powered quiz generator application/` directory:

```env
OPENROUTER_API_KEY=sk-or-v1-your-key-here
```

> **Security:** The `.env` file is listed in `.gitignore` and must never be committed to version control.

### Run

```bash
python -m streamlit run "ai-powered quiz generator application/app.py"
```

The app opens automatically at `http://localhost:8501`.

---

## Usage

1. **Upload** — Drag and drop a `.pptx` file (max 25 MB) onto the upload zone.
2. **Configure** — Choose how many questions (5–30) and select a difficulty level.
3. **Generate** — Click **Generate Questions**. The AI processes slides and builds your quiz (~10–20 seconds).
4. **Quiz** — Answer each question by clicking an option card. Use **Next** / **Back** to navigate.
5. **Results** — View your score, review every question, and read AI explanations for missed answers.
6. **History / Analytics** — Click the sidebar icons to view past sessions and performance stats.

---

## Project Structure

```
ai-powered quiz generator application/
├── app.py              # Main Streamlit application (UI + backend)
├── requirements.txt    # Python dependencies
├── .env                # API key (local only, not committed)
└── README.md           # This file

docs/
├── SPEC.md             # Product specification
└── TECHNICAL_DESIGN.md # Technical architecture and design decisions
```

---

## Dependencies

```
streamlit>=1.35.0
python-pptx>=0.6.23
requests>=2.31.0
python-dotenv>=1.0.0
```

---

## Security Notes

- API keys are loaded from `.env` and never hard-coded.
- The `.env` file is excluded from version control via `.gitignore`.
- No user data is stored persistently — session state is in-memory only.
- All API requests are made server-side; the key is never exposed to the browser.

---

## Academic Context

This application was developed as a homework assignment for **Academic Year 2025–26** as part of the TechVest Agentic AI program. It demonstrates practical use of large language models for educational tooling, prompt engineering for structured JSON output, and custom UI development on top of a Python web framework.

---

## License

For academic and educational use. All rights reserved © 2026 AFIA-REFAL.
