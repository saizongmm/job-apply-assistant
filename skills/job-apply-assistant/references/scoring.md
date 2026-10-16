# Scoring & shortlisting reference

Used in SKILL.md **Step 4** to rank harvested rows before tailoring résumés.

## Signal model (`score_row` in `build_tracker.py`)

| Signal | Points |
|---|---|
| Target keyword appears in `Job Title` or `Company` | **+2** each keyword |
| City matches a user-specified target region (e.g. `Singapore`) | **+1** |
| Seniority in `{Director, Comms, Senior}` | **+1** |

The note column stores `harvest <date>; score=<n>` so the agent can sort.

## Exclusions (drop before tailoring)

- **Recruiters / staffing agencies** — e.g. Robert Walters, Michael Page, Hays,
  Randstad, Manpower, Adecco, PageGroup. These are middlemen, not hiring managers.
- **User-excluded industries** — whatever the candidate says they don't want
  (e.g. tobacco, gambling, fossil fuels).

## Shortlist size

Pick **Top 15–30** by score (user decides the cap). Tailoring is the expensive
step (LLM per job), so keep the batch focused on the highest-signal fits.

## Seniority inference (`infer_role`)

`Director` (director/VP/head/chief/lead) › `Comms` (PR/marcom/公关) ›
`Intern` › `IC` (junior/associate/analyst/graduate) › `Senior` (senior/principal/专家)
› default `Manager`.
