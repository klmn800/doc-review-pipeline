# Document Review Pipeline

A configurable batch document reviewer powered by Claude. Drop Word documents (or Outlook emails containing Word attachments) into a folder, and the pipeline extracts the text, sends each document through a customizable LLM review, and writes structured results to an Excel tracker — including tokens used and cost per document.

The evaluation criteria live entirely in editable prompt files. Swap the prompts and the same pipeline can review contracts, score RFPs, audit policy compliance, screen resumes, or check brand voice — without touching the code.

## Origin

This started as an internal tool for a peer review workflow that needed to flag policy documents missing implementation steps — reviewers were drowning in drafts, and most of the work was triage, not deep analysis. The shipped default prompts still reflect that use case (policy language vs. procedural guidance), but the architecture is deliberately generic: the orchestration code knows nothing about policy. See [`prompts/examples/`](prompts/examples/) for alternative prompt sets demonstrating the same engine doing different jobs.

## Features

- **Folder-watched batch processing** — drop documents in `inbox/`, run the pipeline, get a tracker row per document
- **Email integration** — parses `.msg` (Outlook) and `.eml` files, extracts `.docx` attachments and metadata (sender, subject, summary, due date)
- **List-aware text extraction** — preserves numbered and bulleted formatting from Word documents
- **Claude API integration** — supports Haiku, Sonnet, and Opus with dynamic per-model cost tracking
- **Excel tracker with file-locking handling** — appends rows on each run; if the tracker is open in Excel, retries 3 times then falls back to a timestamped backup so no results are lost
- **Pluggable evaluation prompts** — define your review criteria by editing text files; the orchestration code never needs to change
- **Empty-run hygiene** — scheduled runs that find no work produce a `check_*.txt` log instead of `review_*.txt`, making real review history easy to spot
- **Per-run cost reporting** — every run reports tokens used and dollar cost in the log and console summary

## Project Structure

```
doc-review-pipeline/
├── pipeline.py                # Main pipeline (run this)
├── watcher.py                 # Optional scheduler wrapper - skips run if inbox is empty
├── config.ini.template        # Copy to config.ini and add your API key
├── requirements.txt
├── README.md
├── LICENSE
├── inbox/                     # Drop .docx, .msg, or .eml files here
│   └── processed/             # Auto-created; processed files moved here
├── logs/                      # Auto-created; per-run log files
├── prompts/
│   ├── system_prompt.txt      # The default shipped review prompt
│   ├── instructions/
│   │   └── evaluation_criteria.txt
│   └── examples/              # Alternative prompt sets - copy into prompts/ to swap reviews
│       ├── contract-redline/
│       └── brand-voice/
└── tracker.xlsx               # Auto-created; appended on each run
```

## Installation

```bash
pip install -r requirements.txt
```

Then configure your API key:

```bash
cp config.ini.template config.ini
# edit config.ini and replace `your_key_here` with your Anthropic API key
```

Get an API key at <https://console.anthropic.com/>.

## Usage

### Basic workflow

1. Drop one or more `.docx` files into `inbox/`
2. Run the pipeline:
   ```bash
   python pipeline.py
   ```
3. Check `tracker.xlsx` for the results

### Email workflow

If your review requests come in as emails with attachments, save the email as a `.msg` (Outlook) or `.eml` (universal) file and drop it into `inbox/` instead. The pipeline will:

- Extract all `.docx` attachments
- Parse the email's sender, subject, summary, and due date
- Run each attachment through the review
- Populate the tracker with both the analysis and the email metadata

The default expected email body format:

```
Subject: FOR YOUR REVIEW: <description>   (the prefix is optional)

Project Summary: <one or more paragraphs>
(or "Summary of Draft:", "Summary:", "Draft Summary:")

Comment Due Date: <any reasonable date format - parsed with fuzzy matching>
```

If your emails follow a different convention, edit the regex patterns in `extract_summary` and `extract_due_date` near the top of `pipeline.py`.

### Scheduled runs

`watcher.py` is a lightweight wrapper that checks the inbox before invoking the pipeline. Wire it up to a task scheduler (Windows Task Scheduler, cron, etc.) and it will skip the run when the inbox is empty — keeping your scheduled-run logs clean. See [TASK_SCHEDULER_SETUP.md](TASK_SCHEDULER_SETUP.md) for a Windows walkthrough.

## Customizing the Review

This is the headline feature. The pipeline is domain-agnostic; the prompt files define the review.

To define your own review:

1. **Edit `prompts/system_prompt.txt`** — this is the system message sent to Claude. Tell it what role to play and what JSON response format you want.
2. **Edit or add files in `prompts/instructions/`** — every `.txt` file in this folder is concatenated and appended to the user message. Use this for detailed criteria, examples, or rubrics that would clutter the system prompt.
3. **Run the pipeline as normal.**

The pipeline expects Claude's response to contain a JSON object with `review_status` ("Passed", "Needs Review", or any custom value) and `review_notes` (a brief explanation). If you change the response format, update the parsing in `analyze_document` accordingly.

Example prompt sets live in `prompts/examples/`:

- **contract-redline** — flags documents that are missing standard contract clauses
- **brand-voice** — checks copy against a brand voice guide

To use one, copy its `system_prompt.txt` and `instructions/` folder into `prompts/`, replacing the existing files. The examples are illustrative starting points, not production-ready reviewers — adapt them to your specific criteria.

## Output: tracker.xlsx

A single Excel file appended on each run. Columns:

| Column | Source |
|---|---|
| Document Name | filename |
| Date Processed | run timestamp |
| Review Status | from Claude's response |
| Review Notes | from Claude's response |
| Tokens Used | from API response |
| Cost | computed per model |
| File Path | original document path |
| Due Date | from email (if any) |
| Assigned To | blank — fill manually |
| Summary | from email (if any) |
| Sender | from email (if any) |
| Email Subject | from email (if any) |

If the file is open in Excel when the pipeline runs, the locking detection retries 3 times with 30-second delays. If still locked, results save to a timestamped backup (`tracker_backup_YYYYMMDD_HHMMSS.xlsx`) instead.

## Configuration

`config.ini` controls models, paths, and processing options. The defaults work for a local clone — only the API key is required.

```ini
[API]
anthropic_api_key = your_key_here
model = claude-haiku-4-5-20251001    # or sonnet-4-5, opus-4-5
max_tokens = 2000

[Paths]
input_folder = ./inbox
processed_folder = ./inbox/processed
tracker_file = ./tracker.xlsx
log_folder = ./logs
prompts_folder = ./prompts

[Settings]
preserve_lists = yes
accept_tracked_changes = yes
```

## Cost

Token usage and cost are reported per document and totaled per run. A typical 2-5 page Word document uses ~2,000-5,000 tokens.

Pricing (per million tokens) is hardcoded in `MODEL_PRICING` in `pipeline.py` for Haiku 4.5, Sonnet 4.5, and Opus 4.5. Update the dict if you use a different model.

## Troubleshooting

| Problem | Fix |
|---|---|
| `ERROR: Config file not found` | Run `cp config.ini.template config.ini` and add your API key |
| `ERROR: python-docx not found` | `pip install -r requirements.txt` |
| Documents showing as "Error" | Check the log file in `logs/` — usually a password-protected or corrupted document |
| Tracker has no email metadata | The email body needs a `Project Summary:` (or variant) header and `Comment Due Date:` line. See the email workflow section above. |
| API rate limit errors | The pipeline reports them clearly; wait and re-run. For large batches, process in smaller chunks. |

## License

MIT — see [LICENSE](LICENSE).
