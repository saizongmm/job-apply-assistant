# Job Apply Assistant (通用求职投递助手)

> A reusable [WorkBuddy](https://www.workbuddy.cn) skill that turns "upload your CV + pick a role + pick a region" into a full job-hunting pipeline: **harvest** recent postings, **dedupe** into a tracker, **score & shortlist** the best fits, **AI-tailor** one résumé per job from your real CV, and **open** them in *your own logged-in browser* for manual Apply.

> 中文简介：把"上传简历 + 选岗位 + 选国家/地区"变成一条完整的求职管线——采集近期在招岗位、去重进总表、按匹配度筛 Top、基于你的真实简历用 AI 定制每份简历，并在你本机已登录的浏览器里逐批打开、由你手动点 Apply。

---

## ⚠️ Honest limitations (read first)

This skill is built to be **safe and reliable**, which means two hard no's:

1. **No unattended auto-apply.** LinkedIn and similar sites aggressively anti-bot, and browser session cookies can't be ported across machines/profiles (macOS Keychain cookies are not decryptable cross-path). So the skill **opens** each job in *your* already-logged-in browser and **you** click Apply. This is the only dependable, ban-safe approach.
2. **No-login large-scale scraping hits walls / rate limits.** The WebFetch harvest works for small volumes (dozens–hundreds of fresh postings per role-region per week). At scale you need the user's own signed-in session or a paid jobs API.

Everything else (parsing, dedup, scoring, résumé tailoring) is fully automated.

Résumé tailoring **rewrites/highlights** from your real CV — it never fabricates experience.

---

## Features

- **CV parsing** — `.docx` / `.txt` / `.pdf` → structured `profile.json` (keeps raw text for the LLM).
- **Harvest** — WebFetch LinkedIn search (per role × region), with the "1,000+ lazy-load wall" workaround (qualify the keyword with an industry term).
- **Tracker** — auto-dedup by LinkedIn job-id, city normalization, seniority inference, append-only `application_tracker.xlsx`.
- **Scoring & shortlist** — keyword/city/seniority/industry signal; excludes recruiters and unwanted industries.
- **Per-job tailored résumé** — LLM reads `profile.json` + the JD and emits a `tailored.json`; `make_docx.py` builds a clean `.docx`.
- **Batch opener** — cross-platform (macOS / Linux / Windows), opens only rows that have a tailored résumé and aren't already handled.

---

## Requirements

- Python **3.9+**
- `openpyxl`, `python-docx` (PDF parsing optional: `pdfminer.six` or poppler `pdftotext`)
- Google Chrome installed and **signed in** to the job site on the machine that runs the opener
- A WorkBuddy (or compatible) agent that can run the scripts and use WebFetch

---

## Install

### Option A — via WorkBuddy (recommended for others)

After this repo is pushed to GitHub, anyone can install it by telling WorkBuddy:

> *"Install skills from https://github.com/<your-username>/job-apply-assistant"*

WorkBuddy's `install-github-skills` will clone the repo, run a security audit, and copy `skills/*` into `~/.workbuddy/skills/`. New skills appear on the **next session**.

### Option B — manual

```bash
git clone https://github.com/<your-username>/job-apply-assistant.git
cp -r job-apply-assistant/skills/job-apply-assistant ~/.workbuddy/skills/
```

---

## Quick start

Work in a fresh folder (e.g. `~/apply_developer_sg/`):

```bash
# 0. install deps
pip install openpyxl python-docx

# 1. parse your CV
python3 skills/job-apply-assistant/scripts/parse_cv.py --in cv.docx --out profile.json

# 2. harvest (the agent does this with WebFetch, appending to harvest.tsv)
#    URL template: https://www.linkedin.com/jobs/search/?keywords=<KW>&location=<LOC>&f_TPR=r604800&sortBy=DD
#    If you see "1,000+", qualify the keyword: "Developer Fintech", "Business Development SaaS" ...

# 3. build / append the tracker (dedup by job-id)
python3 skills/job-apply-assistant/scripts/build_tracker.py \
    --tsv harvest.tsv --tracker application_tracker.xlsx --create \
    --wanted "Developer" --cities "Singapore"

# 4. (agent) score -> pick Top N -> for each, write tailored.json from profile.json + JD

# 5. generate tailored résumés
python3 skills/job-apply-assistant/scripts/make_docx.py \
    --profile profile.json --tailored tailored.json \
    --out tailored_resumes/developer_acme.docx

# 6. open a batch in YOUR browser (you click Apply)
python3 skills/job-apply-assistant/scripts/open_jobs.py \
    --tracker application_tracker.xlsx --auto-open --batch 15 --delay 2

# 7. after applying, mark rows
python3 skills/job-apply-assistant/scripts/open_jobs.py --tracker application_tracker.xlsx --done 12 15 23
python3 skills/job-apply-assistant/scripts/open_jobs.py --tracker application_tracker.xlsx --skip 9 31
```

A runnable fictional example lives in [`examples/`](examples/).

---

## Script reference

| Script | Purpose |
|---|---|
| `parse_cv.py` | CV → `profile.json` (docx/txt/pdf) |
| `build_tracker.py` | `harvest.tsv` → deduped `application_tracker.xlsx` (`--create` to init) |
| `make_docx.py` | `profile.json` + `tailored.json` → tailored `.docx` |
| `open_jobs.py` | open / mark tracker rows (cross-platform; macOS/Linux/Windows) |

See [`skills/job-apply-assistant/SKILL.md`](skills/job-apply-assistant/SKILL.md) for the full 7-step agent workflow, scoring details, and the original-pipeline comparison.

---

## Repo layout

```
job-apply-assistant/
├── skills/
│   └── job-apply-assistant/      # <-- the skill (copied into ~/.workbuddy/skills/)
│       ├── SKILL.md
│       ├── scripts/
│       │   ├── parse_cv.py
│       │   ├── build_tracker.py
│       │   ├── make_docx.py
│       │   └── open_jobs.py
│       └── references/
│           └── scoring.md
├── examples/                     # fictional sample CV + harvest + quickstart
├── requirements.txt
├── LICENSE
└── README.md
```

---

## License

MIT — see [LICENSE](LICENSE).
