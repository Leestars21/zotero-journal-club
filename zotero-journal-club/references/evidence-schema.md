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
  "claims": []
}
```

For notes, annotations, figures, equations and claims, record `origin` (`paper`, `user_note`, `presenter_inference`), `locator` (PDF page starting at 1, optional printed page, section, figure/panel, table or equation), `verification` and `destination_slide`. Keep `page_index` only when required by an API and explicitly label it zero-based. A claim may reference multiple sources, and a figure may itself credit another publication; retain both. Never mark an inference as a paper finding or a user note as independently verified.

Metadata can differ between online publication and issue date. Record both when supported. For inaccessible notes or annotations, use `access_status: unavailable`, not a fabricated empty list. For cached content, identity match proves association only; visual and textual checks against the PDF establish content reliability. Public output should include source and figure labels without exposing private notes unnecessarily.
