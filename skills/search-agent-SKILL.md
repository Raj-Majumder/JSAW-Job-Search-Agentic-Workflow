---
name: search-agent
description: Parses LinkedIn and Indeed job-alert emails from Gmail via IMAP and appends new postings to the JobTracker Excel (Job Alerts/Job_Search_-_SOR.xlsx), unfiltered, then labels every row's High Match column from search_config.json. Use this skill to run or resume discovery, or when asked to "find new roles," "check for new postings," or "run the search agent."
---

# Search Agent

First stage of the pipeline. Discovery and labeling only — no scoring against Raj's actual fit, no JD lookup, no writing. Turns job-alert emails into rows in the single source-of-record Excel file.

## Role boundary

This agent does not judge fit against Raj's resume (that's Talent Agent's job), does not fetch full JD text (also Talent Agent's job now), and does not touch the Context Hub. It only discovers postings, appends them as new rows, and applies a cheap keyword label. Everything downstream reads from the same Excel file — this agent never produces separate JSON state.

It also owns the **Applications & Outcomes** tab, but only reactively — it appends a row there when explicitly told a posting was applied to (via `add_application_outcome`, called by JOB-e), never automatically. A drafted cover letter is not the same fact as an application actually sent.

## Inputs

- Gmail via IMAP (not an MCP connector — this pipeline runs standalone)
- `config/search_config.json`: `title_filters` (keywords, case-insensitive substring match), `exclude_keywords`
- `Job Alerts/Job_Search_-_SOR.xlsx` — the JobTracker Excel itself is both an input (read for dedup + last-timestamp) and the output

## Workflow

1. Open the JobTracker workbook. Find the max `Timestamp Captured` across existing rows to know how far back to search Gmail — first-ever run (no data rows yet) falls back to `LOOKBACK_DAYS`.
2. Search Gmail via IMAP for `jobalerts-noreply@linkedin.com` and `donotreply@match.indeed.com` since that timestamp, fetching both plain-text and HTML bodies.
3. Parse postings, **unfiltered** — no title_filters/exclude_keywords applied at this stage:
   - **LinkedIn**: parsed from the HTML body using LinkedIn's real CSS classes (title link: `font-bold` + `text-system-blue-50`; company/location: `text-system-gray-100` + `text-xs`, split on the `·` separator with workplace type in trailing parens; badges: `job-card-flavor__detail`). Plain text drops workplace type entirely and has inconsistent badge wording, so HTML is the only reliable source. A single digest email can contain multiple job cards under one subject line naming only the first job — parse the whole body.
   - **Indeed**: parsed from plain text (single job per email, found via the paragraph containing the `Salary:` line).
4. Deduplicate against existing rows by Job URL. Append genuinely new postings as new rows:
   - Job Title, Company, Location, Job URL, Timestamp Captured always populated.
   - **Job URL is written as a real, clickable Excel hyperlink** (`cell.hyperlink`), not plain text — a project requirement, not a nice-to-have.
   - Hybrid/Remote/On-site and Actively Recruiting? populated where the LinkedIn HTML provided them.
   - Date Posted, Open or Closed, Hiring Team, Pay, Job Description Summary, Match Score, Strong Points, Weak Points all left blank — Talent Agent's job, and only for High Match rows.
5. **Label High Match across every row in the tracker, not just new ones** — this is a post-hoc pass, not a pre-filter. Re-running this step after editing `search_config.json` relabels existing rows too, without needing a fresh Gmail search.
6. Save the workbook.

## Output

Rows in `Job Alerts/Job_Search_-_SOR.xlsx`, sheet "Job Tracker", table `JobTracker`. 18 columns (Serial # outside the table): Job Title, Company, Location, Date Posted, Open or Closed, Hiring Team / Manager, Hybrid / Remote / On-site, Pay / CTC / Salary, Actively Recruiting?, Job Description Summary, Job URL, Timestamp Captured, High Match, Match Score, Strong Points, Weak Points, Cover Created, Resume Updated.

## Error handling

- If Gmail search returns zero matching emails, that's a normal outcome — log it, don't treat as an error.
- If the tracker file itself is missing, raise clearly rather than silently creating a blank one — its formatting, dropdowns, and comments were hand-built and shouldn't be silently replaced.
- Never trust `ws.max_row` for finding where to write — the template has data-validation formatting pre-applied down to row 500, so `max_row` reports 499 even on a genuinely empty tracker. Always scan for the first row with no Job Title.

## Handoff

Downstream: **Talent Agent** reads rows marked `High Match` with no `Job Description Summary` yet — but only when JOB-e dispatches it, either for specific rows or in bulk. This agent never invokes it directly; there's no fixed pipeline sequence anymore, JOB-e decides.
