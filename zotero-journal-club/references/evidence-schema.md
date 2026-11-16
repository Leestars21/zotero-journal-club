# Source ledger contract

The analysis may be held in memory or saved as JSON inside the selected output directory. Preserve these fields when available; use `null` or an explicit access status for unknown information, rather than treating it as an empty result:

```json
{
  "paper_metadata": {"title": null, "authors": [], "year_online": null, "year_issue": null, "doi": null, "origin": null},
  "pdf_source": {"path": null, "attachment_key": null, "page_count": null, "access_status": "unknown"},
  "mineru_content": {"cache_dir": null, "identity_status": "unknown", "markdown": [], "structured_files": [], "reliability": "unchecked"},
  "zotero_notes": [],
  "annotations": [],
  "figures": [],
  "equations": [],
  "references": [],
  "sources": [],
  "claims": []
}
```

For notes, annotations, figures, equations and claims, record `origin` (`paper`, `user_note`, `presenter_inference`), `locator` (PDF page starting at 1, optional printed page, section, figure/panel, table or equation), `verification` and `destination_slide`. Keep `page_index` only when required by an API and explicitly label it zero-based. A claim may reference multiple sources, and a figure may itself credit another publication; retain both. Never mark an inference as a paper finding or a user note as independently verified.

Metadata can differ between online publication and issue date. Record both when supported. For inaccessible notes or annotations, use `access_status: unavailable`, not a fabricated empty list. For cached content, identity match proves association only; visual and textual checks against the PDF establish content reliability. Public output should include source and figure labels without exposing private notes unnecessarily.

For automatic PPT binding checks, save `sources` as a flat ledger with unique `id`, `kind`, `origin`, `locator`, `verification` and optional asset `path`; a source can additionally retain DOI, attachment identity and credits to earlier publications. Set `pdf_source.page_count` when verified so the validator can check one-based page bounds. Use `source_ids` on slides and on individual images/content blocks; use `source_ids` and `destination_slide` on claims. See [slide-plan.md](slide-plan.md) for an executable example. The detailed notes/figures arrays may retain richer material without duplicating private content in the public manifest.

Use `note_kind` such as `human_note`, `human_annotation`, `ai_summary` or `unknown` for Zotero material. This supplements `origin: user_note`; it does not make generated summaries paper evidence. `verification: unchecked` records acquisition only, while `checked_against_pdf` records an actual comparison. A proposed experiment remains `origin: presenter_inference` even when motivated by a verified paper result.
