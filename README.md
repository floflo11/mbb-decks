# MBB Decks

> Generate MBB-style consulting decks with Claude. Action-title storylines, pyramid-principle structure, MECE bullets, navy-and-white editorial visual system. Real `.pptx` output, not Markdown.

## Side-by-side: skill output vs default AI-deck output

Same topic (a data center industry report), same author, same approximate length. Left is what this skill produces. Right is what a generic AI deck tool produces.

**Cover (page 1)**

![Cover comparison](examples/data-center-landscape/compare-cover.png)

**Agenda (page 3)**

![Agenda comparison](examples/data-center-landscape/compare-agenda.png)

**Analytical content (page 11)**

![Content comparison](examples/data-center-landscape/compare-content.png)

**Closing detail (page 15)**

![Closing comparison](examples/data-center-landscape/compare-closing.png)

Full breakdown of the conventions behind each pair lives in [`examples/data-center-landscape/README.md`](examples/data-center-landscape/README.md). Both decks are committed to the repo so you can open them in PowerPoint or Keynote and flip through.

## Why this exists

Generic AI deck tools produce slides that look like AI made them: cluttered layouts, decorative icons, label-style titles ("Market Overview"), and "agenda → conclusion → thank you" structures that no real consultant has ever shipped.

Real MBB decks look different. The story lives in the action titles. The body proves the claim. Charts replace icons. The visual system is restrained on purpose.

This skill encodes those conventions so Claude produces decks that pass the partner-review test on the first draft.

## What's different

- **Ghost deck first.** Claude drafts only the action titles for the entire deck and confirms the storyline with you before expanding any body content. This is how MBB associates actually work and it is the single biggest quality lever.
- **Action-title storyline.** Every slide carries a full-sentence claim, 10 to 15 words, that stands alone. Read top to bottom, the action titles tell the whole argument without the body.
- **Six-section structure.** Cover → Executive Summary → Agenda → Sections (with dividers) → Recommendation (with action table) → Appendix.
- **Chart + commentary on the same slide.** Every chart slide carries a "KEY TAKEAWAYS" panel beside or below it. No bare-chart slides. No "chart on page 5, bullets describing it on page 6."
- **Company logos as bullet markers.** Where a bullet focuses on a single named entity (a company, a regulator, a sovereign fund), the entity's logo replaces the bullet marker. Logos are auto-fetched from the free [Hunter.io logo API](https://hunter.io/changelog/company-logo-api-free/) and cached locally. No API key, no manual sourcing.
- **Visual system.** Georgia headlines, Calibri body, deep navy `#051C2C`, pure black rules, white background. Charts have data labels at 1 decimal on every bar; gridlines and y-axis numbers are suppressed. No em dashes. No date in footer.
- **Citation discipline.** Every numeric claim gets a source line. Footnotes are distinct from sources and prefixed with `*`. Page numbers in the lower right as `Page X / Y`.
- **15-slide cap** for the main deck. Detail goes to the appendix.

## Install

Requires Python 3.9+ and [`python-pptx`](https://python-pptx.readthedocs.io/).

```bash
git clone https://github.com/<your-username>/mbb-decks
cd mbb-decks
pip install python-pptx
```

To use the skill with Claude Code, copy the `mbb-decks` directory into `~/.claude/skills/` (or your project-local `.claude/skills/`).

## Quick start

```bash
# Render the worked example
python scripts/build_deck.py examples/market-entry/input.json examples/market-entry/output.pptx
open examples/market-entry/output.pptx  # macOS
```

To generate your own deck, ask Claude:

> "Build me an MBB-style deck on [topic]. Use the mbb-decks skill."

Claude will produce a ghost deck (action titles only), confirm the storyline with you, then expand the JSON spec and render the `.pptx`.

## Output samples

The repository ships with two worked examples:

- [`examples/market-entry/`](examples/market-entry/): Vietnam market entry recommendation, 14 slides including appendix. Demonstrates cover, executive summary, agenda, three section dividers, content slides with bar charts and two-column layouts (with per-bullet icons), recommendation action table, and appendix backup detail.
- [`examples/data-center-landscape/`](examples/data-center-landscape/): **side-by-side comparison.** Industry report on the data center landscape, 15 slides, rendered both with the skill (`skill-version.pptx`) and with a deliberately generic AI-deck-tool baseline (`vanilla.pptx`). Demonstrates chart-plus-commentary panels, company logos on bullets (AWS, Microsoft, Equinix, Digital Realty, Dominion, Hydro-Quebec, ERCOT, Saudi PIF, etc.), and the below-layout 2-column commentary pattern. The README in that folder breaks down what changes and why it matters.

## How the skill works

| Step | What happens |
|------|--------------|
| 1. Trigger | User asks for an MBB deck, board pre-read, or strategy presentation |
| 2. Ghost deck | Claude drafts action titles for every slide and shows them as a paragraph |
| 3. Confirm | User signs off on the storyline or asks for edits |
| 4. Expand | Claude fills in bullets, charts, footnotes, sources |
| 5. Spec | Claude writes a JSON file with the full deck contents |
| 6. Render | The script generates the `.pptx` |

See [SKILL.md](SKILL.md) for the full workflow Claude follows.

## Roadmap

This is the first in a planned series of "house style" skills covering the four office artifacts that consultants and bankers ship every day.

- **`mbb-decks`** (this repo): slide decks
- `mbb-models`: financial models (three-statement, LBO, sensitivity)
- `mbb-memos`: Word documents (IC memos, board prereads, one-pagers)
- `mbb-mail`: partner-style email and follow-ups

Subscribe at [Substack link TBD] for new release announcements and a weekly post on a single MBB convention rebuilt from a famous public deck.

## About the author

Built by [Iris Meng](https://www.linkedin.com/in/yilin-meng/). Three years as a Senior Associate at EY Transaction Economics in New York, plus senior management office work at listed companies where she drafted decks alongside ex-MBB executives. This skill codifies the conventions she has spent six years working against.

## Contributing

Pull requests welcome, especially for additional chart families, more worked examples, and theme variations. Please open an issue first for major changes so we can align on direction.

## License

MIT. See [LICENSE](LICENSE).
