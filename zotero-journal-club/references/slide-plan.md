# Authored slide plans and source bindings

Use this reference when invoking the optional Python helper. It renders an authored plan; it does not read the paper, choose conclusions or verify them scientifically.

## Common fields

For a finished untemplated academic talk, add `"design": {"preset": "academic-blue", "label": "组会文献汇报"}` to the top-level plan. The preset uses editable chapter/page headings, local captions, structured evidence and a source footer. The legacy mode without `design` is retained for compatibility and structural drafts. Read [visual-quality.md](visual-quality.md) for reference-based planning and acceptance. Do not combine an independent preset with `--template`; use inspected native template layouts instead.

```json
{
  "duration_minutes": 15,
  "slides": [{
    "id": "result-1",
    "title": "A supported result",
    "bullets": ["Brief context"],
    "source": "Author et al., year, Figure 3, PDF p. 5",
    "source_ids": ["F3"],
    "notes": "Explain the evidence and transition to the next slide.",
    "duration_seconds": 90,
    "reading_questions": ["Does the comparison isolate the proposed change?"],
    "figure_reading": {
      "comparison": "Method A and B",
      "conditions": "Same dataset and training budget",
      "observation": "Describe the measured difference",
      "claim": "State only what the evidence supports",
      "scope": "The evaluated dataset and split"
    },
    "takeaways": ["A transferable design choice and its validation condition"]
  }]
}
```

Every slide needs `title`, a nonempty human-readable `source` label and substantive `notes`. Structured reading guidance and duration are appended to speaker notes. `id` defaults to `slide-001`, etc.; assign unique explicit IDs when using claim destinations. The time-budget check warns when supplied per-slide durations exceed the talk budget. It does not infer durations for missing entries.

## Content types

The academic preset also accepts `page_type: cover` with `subtitle` and a `metadata` list, `page_type: agenda` with two to four `agenda` titles, and `content`/`closing` pages with a `section`, concise `takeaway` and an `emphasis` list of literal phrases to highlight. These remain native editable objects. Keep metadata and emphasis specific to the new talk.

For related figure/method regions, use one to three titled panels. Put all content inside the panels instead of combining `panels` with slide-level bullets/images/table/flow/equation/columns:

```json
"panels": [
  {"title": "Original evidence", "images": [{"path": "assets/figure-3.png", "caption": "Figure 3", "source_ids": ["F3"]}]},
  {"title": "Interpretation", "bullets": ["A supported observation"], "source_ids": ["F3"]}
]
```

For concise text comparisons and synthesis, the academic preset supports two to four `summary_rows` with editable numbered titles, details and rules. This avoids stretching short text into oversized empty panels. Do not mix it with other slide-level content:

```json
"summary_rows": [
  {"title": "Comparison", "text": "Same data and training budget", "source_ids": ["SETUP"]},
  {"title": "Scope", "text": "The evaluated conditions", "source_ids": ["RESULT"]}
]
```

Panel/row source IDs are included in source-binding checks. The checker also verifies cover metadata, outline entries, takeaways, local captions, row text and the speaking duration retained in notes. The academic preset lays out native flow numbers/titles/details separately, places equation variables beside the model where space allows, and uses table row heights based on content. It rejects obviously crowded text or table cells. A successful build still needs full visual review.

Ordinary `bullets`, legacy `image` and one/two `columns` remain supported. Columns take precedence over ordinary bullets and work with or without a template. Legacy `image_focus: true` reserves a larger area for a single figure. For richer content, choose one of `images`, `table`, `flow` or `equation` per page; combining them with legacy `image` or `columns` is rejected. Brief `bullets` may introduce a rich page.

Captioned images (one to four; `image_columns` is 1 or 2):

```json
"images": [
  {"path": "assets/figure-3a.png", "caption": "Figure 3a — caption", "source_ids": ["F3A"]},
  {"path": "assets/figure-3b.png", "caption": "Figure 3b — caption", "source_ids": ["F3B"]}
]
```

Native editable table:

```json
"table": {
  "headers": ["Method", "Metric (unit)"],
  "rows": [["A", "0.12 ± 0.02"], ["B", "0.09 ± 0.01"]],
  "caption": "An explicitly identified excerpt; values must be checked against the source",
  "font_pt": 16,
  "source_ids": ["T1"]
}
```

Native process diagram (two to six steps; longer processes use numbered rows):

```json
"flow": {
  "steps": [{"title": "Data", "detail": "Measured response"}, "Fit", "Validate"],
  "source_ids": ["METHOD"]
}
```

Editable simple equation and variables:

```json
"equation": {
  "text": "y = ax + b",
  "variables": [{"symbol": "a", "meaning": "Slope and its unit"}, "b: intercept"],
  "source_ids": ["E1"]
}
```

Equation text is plain Unicode, not LaTeX parsing or an Office native mathematical equation. Use a verified source crop for complex notation. Native tables/flows remain editable; original scientific plots are preserved images.

The academic preset accepts a `caption` in the `flow` object to label its source or meaning. Use `column_titles` with academic `columns` when their meanings need labels. Single-image academic pages may use `caption`; multi-image pages give each image its own caption. Choose a layout according to the evidence rather than relying on one repeated arrangement.

For templates, inspect the real layout first. Ordinary content can use existing placeholders. Rich content uses the first suitable body placeholder, or the explicitly supplied `regions.title` and `regions.body` on a custom layout; `body` is the whole content area, not a half-width text column. Existing title/body font controls apply where supported. Review every rendered page for fit and template artwork.

## Evidence sidecar

```json
{
  "pdf_source": {"page_count": 12},
  "sources": [{
    "id": "F3",
    "kind": "figure",
    "origin": "paper",
    "path": "assets/figure-3.png",
    "locator": {"pdf_page": 5, "figure": "Figure 3"},
    "verification": "checked_against_pdf"
  }],
  "claims": [{
    "id": "C1",
    "text": "A supported finding",
    "origin": "paper",
    "source_ids": ["F3"],
    "destination_slide": "result-1"
  }]
}
```

Use one consistent set of IDs in the real plan; the snippets illustrate separate content types. Relative image paths are resolved against the plan file, and ledger asset paths against the ledger file. Copying identical source assets is supported because binding checks compare image bytes. Slides and content blocks inherit slide-level source IDs when appropriate; individual image IDs are recommended for precise attribution. A claim may target a slide ID, one-based slide number or a list of those destinations. A paper finding must include paper evidence, not only an AI/human note or presenter inference.

## Commands

```text
python -X utf8 scripts/build_ppt.py slide_plan.json presentation.pptx --template template.pptx
python -X utf8 scripts/validate_ppt.py presentation.pptx --plan slide_plan.json --evidence evidence.json
python -X utf8 scripts/extract_figures.py source.pdf regions.json assets --dpi 220
```

Omit `--template` for the academic default. The build and crop helpers refuse to overwrite existing outputs. The crop plan is a list of `{id, pdf_page, bbox, figure/panel/table/equation}` entries; `bbox` uses observed PDF-point coordinates. Generated `figure_manifest.json` records acquisition and file identity, with verification still `unchecked` until actual PDF comparison. Build a new output when revising a talk, and keep plans, evidence and visual QA with it.
