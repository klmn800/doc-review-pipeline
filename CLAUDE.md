# CLAUDE.md — doc-review-pipeline

A configurable batch document reviewer powered by Claude. The portfolio `CLAUDE.md` and
`../PORTFOLIO_STANDARD.md` also apply (they load in every session under `sandbox_projects`).

## Read first

- `PROJECT.md`: status block, decisions, open questions.
- `README.md`: what the tool does, how it's configured and run.

## Rules

- **Real documents are sensitive.** The inbox may hold DTA policy drafts and emails. Don't
  send real work documents through the pipeline unless the AI service has been approved for
  them; use made-up test documents otherwise. Never commit anything from `inbox/`,
  `processed/`, `logs/` or the tracker (the `.gitignore` covers these).
- **No secrets in files that are tracked.** The API key goes in `config.ini` (ignored) or the
  `ANTHROPIC_API_KEY` environment variable.
- Evaluation logic lives in `prompts/`; the orchestration code stays generic.
