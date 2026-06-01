# Access Model Operating System

> Coursework submission (5/30 cohort) — Mastering Agentic AI.

An interactive, game-like training tool for Future 2 staff (educators called
"Guides") to learn, practice, and implement the **Access Model**. It is built as a
modern, SaaS-style web app rather than a document browser: a dark navigation rail,
gradient dashboard, white cards with hover lift, stat tiles, and decision-based
missions.

## What it does

- **Learn (decision-based).** Nine **Missions** form a three-level *Future 2 Guide
  Certification* pathway. Each mission shows one real scenario and one decision at a
  time. A correct answer is required to complete a mission; missions unlock linearly.
- **Practice.** A **Scenario Challenge** with a shuffled deck of the full scenario
  bank, grade-band and category filters, three answer choices, instant coaching, and
  live accuracy metrics.
- **Implement.** Card-based toolkits: Repair Protocols, Launch Toolkit, Check Charts,
  and Alpha to Access.
- **Reference.** Foundations (the thinking behind the model) and the Guide Role
  (the seven principles through what a Guide looks for).
- **Progress.** A dashboard home, plus Missions Completed and Scenarios Completed
  history pages. All progress lives in session state only (nothing is persisted).

## Training quality: distractors and the audit guardrail

The two wrong answers for every scenario are generated to model realistic,
well-intentioned implementation **mistakes**: one *over-functioning* (doing too much
/ solving for the student) and one *under-functioning* (stepping back too far /
ignoring the issue), keyed to the scenario's Access Model principle.

A banned-pattern audit (`audit_scenarios`) verifies that **no** generated answer
choice across all scenarios contains punitive, exclusionary, shaming,
compliance-first, or power-struggle language. The result is recorded in session
state and surfaced on the dashboard if it ever fails.

## Data source

The Excel workbook `access_model_training_data_v4.xlsx` at the repo root is the only
data source. It is loaded once with
`pandas.read_excel(path, sheet_name=None, engine="openpyxl")` behind
`@st.cache_data`. If the workbook is missing, the app regenerates a representative
one matching the required schema (seven sheets: Scenarios, Access Model Principles,
Launch Lessons, Check Chart Alignments, Alpha to Access, Brainlift References, Data
Dictionary).

Grade bands are always rendered as "Grades 3-4 / 5-6 / 7-8", "All grades" (for
`All`), and "Staff" (for `Adult`). The word "ages" is never used.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

To (re)build the workbook without launching the app:

```bash
python app.py build
```

## Run the smoke test

An automated smoke test using Streamlit's `AppTest` exercises every page, the
mission unlock and correct-answer gate, the Scenario Challenge, and the distractor
audit:

```bash
python -m pytest test_app.py -q
# or
python test_app.py
```

## Deploy on Streamlit Community Cloud

Point a new app at this repository with `app.py` as the entry point. The only
dependencies are Streamlit, pandas, and openpyxl (see `requirements.txt`). No login,
database, network assets, or external services are used.

## Files

- `app.py` — the entire application (single file).
- `access_model_training_data_v4.xlsx` — the data source workbook.
- `requirements.txt` — dependencies.
- `test_app.py` — AppTest smoke test.
