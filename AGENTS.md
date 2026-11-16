# Repository guidance

This repository contains one installable skill in `zotero-journal-club/`. Preserve this layout for existing installations.

- Keep the repository root for README mirrors, LICENSE, `.gitignore` and this file.
- Keep skill code and reference guidance in `zotero-journal-club/`, regression tests in `tests/`, and maintained synthetic examples in `examples/`.
- Use `zotero-journal-club/scripts/render_preview.mjs` for compatible previews; do not add duplicate rendering helpers in tests.
- Keep private papers, notes, reference presentations and generated talks in ignored `output/`, or their existing user-owned locations. Publish only abstract guidance, synthetic examples and concise QA records.
- Use one `output/<paper-id>/` directory per paper. Retain the final PPTX, analysis, source manifest, QA, slide/evidence sidecars, required assets and one preview set.
- Put task-created drafts, temporary scripts and repeated renders in `output/<paper-id>/.work/`. After verifying the final package, remove these reproducible intermediates; preserve user-owned original files and final deliverables.
- Archive development notes worth retaining in `output/archive/`; do not leave ad hoc reports at the root.
- Synchronize Chinese and English README files when directory policy or workflow changes. Run focused checks for cleanup and the relevant regression tests for behavior changes.
