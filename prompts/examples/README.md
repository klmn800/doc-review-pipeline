# Example Prompt Sets

These are **illustrative starting points** for using the pipeline on different document types. They are not production-ready reviewers — adapt them to your specific criteria, tolerances, and edge cases before relying on them.

Each example folder mirrors the structure of the top-level `prompts/`:

```
<example-name>/
├── system_prompt.txt
└── instructions/
    └── criteria.txt
```

## Using an example

To swap the active review, copy an example's contents into the top-level `prompts/` directory, replacing the existing files:

```bash
# from the repo root
cp prompts/examples/contract-redline/system_prompt.txt prompts/system_prompt.txt
cp -r prompts/examples/contract-redline/instructions/. prompts/instructions/
```

Then run `python pipeline.py` as usual.

## Available examples

- **`contract-redline/`** — flags contracts missing standard clauses (parties, term, governing law, termination, indemnification, payment terms, signatures)
- **`brand-voice/`** — checks copy against a brand voice description; flags drafts that drift from the defined tone

## Writing your own

The pipeline expects Claude to return JSON with two fields:

```json
{
  "review_status": "Passed",
  "review_notes": "Brief explanation."
}
```

`review_status` can be any string you choose ("Passed", "Needs Review", "Approved", "Rejected", "Score: 7/10" — whatever fits your workflow). The pipeline writes it directly to the tracker.

Whatever criteria you use, keep `review_notes` short — it has to fit in an Excel cell that a human will skim.
