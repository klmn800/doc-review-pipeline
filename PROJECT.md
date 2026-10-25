# doc-review-pipeline — Project Status

## At a glance

*Updated 2026-10-25 (first written by the portfolio level when the repo was cloned into
`sandbox_projects`; this project's sessions own it from here)*

### Current state
**Parked (Ben, 2026-10-25).** An early prototype, kept as a reference for when this work
resumes, not as a base to extend as-is. It batch-reviews Word documents (or `.msg`/`.eml`
emails with `.docx` attachments) with Claude against editable prompts and writes results and
cost to an Excel tracker. The original plan was Power Automate with a Copilot step; there's no
Copilot at work, and no approved endpoint for a Claude step either, so it stalled. It predates
Ben's current standards. Not installed or run on this laptop.

### Needs Ben
Nothing.

### Blockers
None.

### Next
Nothing until the EOHHS sandbox is understood (`shared-knowledge/eohhs-sandbox`); then it's on
Ben's queue. When it resumes, the first step is to review it against the current standards and
decide what to keep (prompt-file design, email parsing and the tracker's lock fallback look
reusable) and what to redo.
