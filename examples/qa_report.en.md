# Validation record

[中文](qa_report.md) | English

Validated on November 16–17, 2026. Coverage includes interface/data contracts, a real paper, template compatibility and finished-deck visual requirements. Full papers, private notes, the user's reference deck, original figure crops and validation talks remain in ignored local `output/`.

## Automated and synthetic checks

All 50 tests passed. They cover MCP envelopes/errors, pagination and ambiguous identities, MinerU attachment association, PDF crop provenance and output protection, default/template generation, image/evidence bindings, native tables/flows, notes and durations, academic layouts and numbered synthesis rows. Removing planned text, substituting a source image or deleting the duration from notes produces a validation failure. The Office preflight test confirms existing output is preserved before Office starts.

- General object example: 4 slides, 22 editable-text counts and 1 native table; no content/evidence issues or warnings.
- Academic example: 6 slides, 85 editable-text counts and 1 native table; no content/evidence issues or warnings. All pages were rendered with artifact-tool and visually reviewed. The public preview contains only synthetic material.
- Editable-text counts include table cells. Structure checks are distinct from scientific and visual review.

## Real Zotero paper and template

The test paper was Li, Miranda and Pinho, *Probabilistic failure prediction for early-stage design of novel composite wing structures via heteroscedastic Bayesian learning*, DOI [10.1016/j.ast.2026.113816](https://doi.org/10.1016/j.ast.2026.113816). The online year is 2026 and issue year is 2027.

Read-only access resolved the parent, exact PDF attachment, associated MinerU cache and one AI reading summary. The summary guided reading; publication metadata was corrected against the PDF and item metadata. The real annotation query was empty. Figure 9/12 requests returned no images, so inspected original PDF regions were cropped locally with one-based pages, coordinates, PDF hashes and 300 DPI assets.

Earlier default and user-template versions each contain 16 slides, 8 pictures and 1 native table; all 32 V2 pages were visually inspected. The finished academic version contains 17 slides, 235 editable-text counts, 8 original figures/panels and 1 native table, with notes and source labels throughout. Final content/evidence checks have no issues; small-picture warnings were reviewed page by page. The complete Figure 9 provides context, followed by two slides magnifying its four panels. Table 3 preserves MAE, means and standard deviations and states that metrics are in PCA coefficient space.

## Finished reference and visual improvements

Every page of a supplied 21-slide Chinese academic deck was inspected. Its chapter/local-title hierarchy, blue styling, restrained red emphasis, aligned evidence/captions, method relationships, tables, numbered synthesis and takeaways informed the baseline. Private names, research results, institutional pictures and the full reference deck were excluded from the public repository. A local `quality_reference` can retain the visual benchmark independently of the build template.

Actual previews led to fixes for clipped multi-line introductions, inherited white table-header text, oversized empty text panels, weak flow hierarchy, scattered equation explanations and overly tall table rows. All 17 final paper slides and 6 synthetic slides were inspected; structure/parameter, method, table, original-figure and synthesis pages received full-size review. Axes, legends, units and panel labels remain intact and input files were preserved. The [visual quality contract](../zotero-journal-club/references/visual-quality.md) makes reference comparisons part of finished-deck acceptance.

## Checks not completed

- No real annotation content was available, so real annotation integration was not exercised; the retrieved note was an AI summary.
- Native PowerPoint PDF export did not complete because the existing instance was inaccessible through COM (`0x800401E3`). The user's app was neither terminated nor replaced. Visual review used artifact-tool; target-app fonts/effects can differ.
- The reference produced 21 complete PNG files that were opened individually, but its compatible renderer returned nonzero after export without further diagnostics. Rendering of the final 17-slide paper and 6-slide synthetic decks completed normally.
- Complex mathematics still requires verified original crops or reliable mathematical objects; editable helper equations are simple Unicode text.
