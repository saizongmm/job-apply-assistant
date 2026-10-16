# Quickstart (fictional example)

This walks the skill end-to-end with the bundled fictional candidate **JANE DOE**
(a Singapore backend engineer) — no real CV or LinkedIn session required.

```bash
# from the repo root
pip install -r requirements.txt

# 1. parse the sample CV
python3 skills/job-apply-assistant/scripts/parse_cv.py \
    --in examples/sample_cv.txt --out /tmp/jaa_demo/profile.json

# 2. build the tracker from the sample harvest (note the recruiter row gets a lower signal)
python3 skills/job-apply-assistant/scripts/build_tracker.py \
    --tsv examples/sample_harvest.tsv \
    --tracker /tmp/jaa_demo/application_tracker.xlsx --create \
    --wanted "Developer,Backend,Go" --cities "Singapore"

# 3. (agent step) read profile.json + each JD, write tailored.json, then:
python3 skills/job-apply-assistant/scripts/make_docx.py \
    --profile /tmp/jaa_demo/profile.json \
    --tailored /tmp/jaa_demo/tailored.json \
    --out /tmp/jaa_demo/tailored_resumes/jane_acme.docx

# 4. dry-run the opener (lists rows that have a résumé + aren't handled)
#    NOTE: the agent also writes the résumé filename into the tracker's
#    "Tailored Resume" column (only that column + Notes) in Step 5, which is
#    what lets the opener pick the row up. Without it, the dry-run shows 0.
python3 skills/job-apply-assistant/scripts/open_jobs.py \
    --tracker /tmp/jaa_demo/application_tracker.xlsx --dry-run
```

The `make_docx.py` step needs a `tailored.json` written by the LLM. Shape:

```json
{
  "target": {"title": "Backend Developer", "company": "Fintech Co", "location": "Singapore"},
  "summary": [
    "6 yrs backend (Go/Python) on payments & identity platforms at 4M txn/day scale.",
    "Owns event-driven microservices (Kafka) and low-latency gRPC APIs.",
    "Cut cloud spend 28% via autoscaling and PostgreSQL tuning."
  ],
  "experience": [],
  "education": []
}
```

Open the generated `application_tracker.xlsx` to see the deduped, scored table.
