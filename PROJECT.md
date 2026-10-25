# doc-review-pipeline — Project Status

## At a glance

*Updated 2026-10-25 (first written by the portfolio level when the repo was cloned into
`sandbox_projects`; this project's sessions own it from here)*

### Current state
A working prototype: batch-reviews Word documents (or `.msg`/`.eml` emails with `.docx`
attachments) dropped in `inbox/`, sends each one to Claude with editable prompts, and writes
results and cost to an Excel tracker. Built for a DTA policy peer-review workflow; it stalled
because no approved API endpoint was available at work. Cloned onto the work laptop 2026-10-25.
Not yet installed or run here.

### Needs Ben
- Is this active again, and for what? If it's to run on work documents, it needs an approved
  AI service (same open question as db-chat: the EOHHS sandbox's Bedrock catalog). Until then,
  no real DTA documents go through it on a personal API key.

### Blockers
None.

### Next
Decide its future (above). If active: install on Python 3.14, run it against made-up test
documents, and bring `config.ini.template`'s model list up to date (it names 4.5-era models).
