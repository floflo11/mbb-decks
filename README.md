# MBB Decks

> Generate MBB-style consulting decks with Claude. Action-title storylines, pyramid-principle structure, MECE bullets, navy-and-white editorial visual system. Real `.pptx` output, not Markdown.

## With skill vs without skill

Same prompt, same topic (a data center industry report), same author. **Left** is what Claude produces with this skill installed. **Right** is what Claude produces without it (the default AI-deck-tool aesthetic).

![Cover comparison](examples/data-center-landscape/compare-cover.png)

![Agenda comparison](examples/data-center-landscape/compare-agenda.png)

![Content comparison](examples/data-center-landscape/compare-content.png)

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

## Prerequisites

- **Python 3.9+** with [`python-pptx`](https://python-pptx.readthedocs.io/) installed.
- **[Claude Code](https://claude.com/claude-code)** if you want Claude to drive the workflow end to end. You can also use the skill manually by hand-writing a JSON spec and running the renderer yourself; see "Use it without Claude" below.
- **Internet access on first render**, so the script can auto-fetch any logo it does not find in `assets/logos/`. The 17 logos used in the worked examples are pre-cached and committed, so the bundled examples render fully offline.

## Install

```bash
git clone https://github.com/floflo11/mbb-decks
cd mbb-decks
pip install python-pptx
```

To make the skill available to Claude Code:

```bash
# Globally (every project)
cp -R . ~/.claude/skills/mbb-decks

# Or project-local (just the current repo)
mkdir -p .claude/skills && cp -R /path/to/mbb-decks .claude/skills/
```

## Quick start

```bash
mkdir -p out

# 1. Render the side-by-side comparison so you can flip through it
python scripts/build_deck.py examples/data-center-landscape/input.json out/data-center.pptx
python examples/data-center-landscape/build_vanilla.py out/data-center-vanilla.pptx
open out/data-center.pptx out/data-center-vanilla.pptx  # macOS

# 2. Render the market-entry example
python scripts/build_deck.py examples/market-entry/input.json out/market-entry.pptx
open out/market-entry.pptx
```

## Generate your own deck

The skill is designed to be driven by Claude Code, but the JSON-plus-script split means you can run it any way you like.

### Driven by Claude (recommended)

Open Claude Code in any directory with the skill installed and prompt:

> Build me an MBB-style deck on [your topic]. Use the mbb-decks skill.

Claude follows the workflow encoded in `SKILL.md`:

1. **Ghost deck.** Drafts only the action titles for every slide and reads them back to you as a paragraph.
2. **Confirm.** You sign off on the storyline, ask for edits, or change the structure.
3. **Expand.** Claude fills in bullets, charts, footnotes, sources, and selects company logos for entity-anchored bullets.
4. **Render.** Claude writes the JSON spec to disk and runs `scripts/build_deck.py` to produce the `.pptx`.

### Use it without Claude

Hand-write a JSON spec following the schema in [`SKILL.md`](SKILL.md) (use `examples/data-center-landscape/input.json` as a template), then run:

```bash
python scripts/build_deck.py path/to/your-spec.json out/your-deck.pptx
```

To pre-warm the logo cache for a list of companies before rendering:

```bash
python scripts/download_logos.py acme.com partner.com regulator.gov
```

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| Bullets render with `•` instead of a logo | Logo fetch failed (no internet, firewall, or unknown domain) | Run `scripts/download_logos.py <domain>` manually, or check `https://logos.hunter.io/<domain>` returns a PNG in your browser |
| PowerPoint shows "needs to repair" on open | Stale build with float EMU values | Re-render with the latest `build_deck.py`; integer-EMU coercion is now everywhere it needs to be |
| Headline overflows the right edge | Action title exceeds 110 characters at 18pt Georgia | Shorten the headline; the build script prints a warning naming the slide |
| Chart bars have no value labels | Old chart spec without `unit`, or a chart family the renderer does not yet style | Bar charts (`content_type: "bar_chart"`) are fully styled; other families fall back to defaults |

## Output samples

The repository ships with two worked examples. Both have a committed `.pptx` in their folder so you can preview without running anything.

- [`examples/data-center-landscape/`](examples/data-center-landscape/): industry report on the data center landscape, 15 slides, rendered both with the skill (`skill-version.pptx`) and with a deliberately generic AI-deck-tool baseline (`vanilla.pptx`). This is the deck behind the comparison images at the top. Demonstrates chart-plus-commentary panels, company logos on bullets (AWS, Microsoft, Equinix, Digital Realty, Dominion, Hydro-Quebec, ERCOT, Saudi PIF, etc.), and the below-layout 2-column commentary pattern. The folder README breaks down what changes and why it matters.
- [`examples/market-entry/`](examples/market-entry/): Vietnam joint-venture recommendation, 14 slides including appendix. Demonstrates cover, executive summary, agenda, three section dividers, content slides with bar charts and two-column layouts (with per-bullet icons), recommendation action table, and appendix backup detail.

See [SKILL.md](SKILL.md) for the full workflow Claude follows, including the JSON schema, action-title rules, MECE-bullet conventions, and the chart-plus-commentary patterns.

## Roadmap

This is the first in a planned series of "house style" skills covering the four office artifacts that consultants and bankers ship every day.

- **`mbb-decks`** (this repo): slide decks
- `mbb-models`: financial models (three-statement, LBO, sensitivity)
- `mbb-memos`: Word documents (IC memos, board prereads, one-pagers)
- `mbb-mail`: partner-style email and follow-ups

A Substack covering one MBB convention per post, rebuilt from a famous public deck, is in the works. Watch this repo or follow the author on LinkedIn (below) for the launch.

## About the author

Built by [Iris Meng](https://www.linkedin.com/in/yilin-meng/), co-founder of New York AI Labs, working at the intersection of finance and AI. Previously: three years as a Senior Associate at EY Transaction Economics in New York, plus senior management office work at listed companies where she drafted decks alongside ex-MBB executives. This skill codifies the consulting and banking deck conventions she has spent six years working against.

## Contributing

Pull requests welcome, especially for additional chart families, more worked examples, and theme variations. Please open an issue first for major changes so we can align on direction.

## License

MIT. See [LICENSE](LICENSE).
