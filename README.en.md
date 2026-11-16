# Zotero Journal Club

[中文](README.md) | English

Project: [Leestars21/zotero-journal-club](https://github.com/Leestars21/zotero-journal-club)

Create editable Chinese journal-club presentations from a Zotero paper or local PDF, with paper analysis, source mapping, speaker notes and QA. The workflow prioritizes existing Zotero/MinerU materials and the user's PowerPoint template.

## Install and requirements

Ask Codex to install it:

```text
Install the zotero-journal-club Skill from https://github.com/Leestars21/zotero-journal-club.
```

Alternatively, download the repository and copy its `zotero-journal-club/` subdirectory to `~/.codex/skills/zotero-journal-club/` (Windows: `%USERPROFILE%\.codex\skills\zotero-journal-club\`). Start a new session and invoke `$zotero-journal-club`.

Codex and presentation authoring tools are required. The build and validation scripts need Python 3, `python-pptx` and Pillow; the PDF inspection helper needs PyMuPDF. Zotero access requires a suitable read-only interface, with a local PDF as fallback. PowerPoint or an existing artifact-tool runtime can render previews; their display may differ. No additional paid model API is required.

## Usage

```text
$zotero-journal-club Prepare a 15-minute Chinese group presentation on the paper in Zotero titled ...
$zotero-journal-club Prepare a talk from this local PDF using my PPTX template.
```

DOIs, item keys and links are also accepted. Supply a template in the request or configure `default_template` in `~/.codex/zotero-journal-club.json` using the [example](zotero-journal-club/config.example.json). `ZOTERO_JOURNAL_CLUB_CONFIG` selects another configuration file. Explicit templates take precedence, and missing selected files cause an error. Codex performs the research and verification; scripts do not understand papers. The optional builder fills placeholders or explicitly specified regions on a custom layout after inspection. It retains theme and masters; complex templates need manual adjustment and visual review. Automated checks cannot establish accuracy or catch every overflow.

## Validation and feedback

Run `python -m unittest discover -s tests` and the Codex `quick_validate.py` skill validator. Report problems through [GitHub issues](https://github.com/Leestars21/zotero-journal-club/issues), including the input type, OS, error and expected behavior; do not publish private notes, full papers or Zotero databases.

## License

See [LICENSE](LICENSE).
