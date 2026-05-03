# MBB Decks

> Real consulting decks from Claude. Not slop.

Action-title storylines, MECE bullets, chart-plus-commentary panels, company logos on bullets, real `.pptx` output. Built by [Iris Meng](https://www.linkedin.com/in/yilin-meng/), co-founder of New York AI Labs, working at the intersection of finance and AI.

## With skill vs without skill

Same prompt, same topic (a data center industry report), same author. **Left** is what Claude produces with this skill installed. **Right** is what Claude produces without it (the default AI-deck-tool aesthetic).

![Cover comparison](examples/data-center-landscape/compare-cover.png)

![Agenda comparison](examples/data-center-landscape/compare-agenda.png)

![Content comparison](examples/data-center-landscape/compare-content.png)

![Closing comparison](examples/data-center-landscape/compare-closing.png)

Full breakdown of the conventions behind each pair lives in [`examples/data-center-landscape/README.md`](examples/data-center-landscape/README.md). Both decks are committed to the repo so you can open them in PowerPoint or Keynote and flip through.

## Install

This skill is designed for [Claude Code](https://claude.com/claude-code), Anthropic's command-line agent. If you don't have it yet, install it from <https://claude.com/claude-code> first.

### Step 1: Install the renderer dependencies

The skill writes real `.pptx` files using [`python-pptx`](https://python-pptx.readthedocs.io/). You need Python 3.9 or newer.

```bash
python3 --version          # check you have 3.9+
pip install python-pptx    # the renderer engine
```

### Step 2: Add the skill to Claude Code

Open any project folder in Claude Code (run `claude` in your terminal). Inside the prompt, paste both lines:

```
/plugin marketplace add floflo11/mbb-decks
/plugin install mbb-decks
```

The first line pulls this repo as a plugin source. The second installs the skill. Two commands is the minimum Claude Code requires for any plugin install; there is no shorter single-command path today.

### Step 3: Verify

In a regular terminal:

```bash
ls ~/.claude/plugins/cache/
# you should see a mbb-decks directory
```

### Step 4: Use it

In any Claude Code session, prompt:

> Build me an MBB-style deck on [your topic]. Use the mbb-decks skill.

Claude follows the workflow encoded in `SKILL.md`: drafts a ghost deck (action titles only), confirms the storyline with you, expands the JSON spec, and renders the `.pptx`. Output lands in your current working directory.

### Updating or removing

Inside Claude Code:

```
/plugin update mbb-decks      # pull latest version
/plugin uninstall mbb-decks   # remove the skill
```

## Skip Claude Code: render decks from a JSON spec

If you don't use Claude Code and just want the renderer, clone the repo and run the build script directly with a hand-written spec:

```bash
git clone https://github.com/floflo11/mbb-decks
cd mbb-decks
pip install python-pptx
mkdir -p out

# Render the bundled examples
python scripts/build_deck.py examples/data-center-landscape/input.json out/data-center.pptx
python examples/data-center-landscape/build_vanilla.py out/data-center-vanilla.pptx
open out/data-center.pptx out/data-center-vanilla.pptx  # macOS
```

Use [`examples/data-center-landscape/input.json`](examples/data-center-landscape/input.json) as a template and follow the schema in [`SKILL.md`](SKILL.md) to write your own spec.

## Why this exists

Generic AI deck tools produce label titles, bullet soup, decorative icons, pie charts, and a "Thank You" closer. Real consulting decks tell a story through action titles, prove each claim with one chart or three MECE bullets, and end on a recommendation table with named owners. This skill encodes those conventions so Claude produces decks that pass the partner-review test on the first draft.

## What's different

- **Ghost deck first.** Claude drafts only the action titles for the entire deck and confirms the storyline with you before expanding any body content. The single biggest quality lever.
- **Action-title storyline.** Every slide carries a full-sentence claim, 10 to 15 words, that stands alone. Read top to bottom, the action titles tell the whole argument.
- **Chart + commentary on the same slide.** Every chart slide carries a "KEY TAKEAWAYS" panel beside or below it. No bare-chart slides. No "chart on page 5, bullets describing it on page 6."
- **Company logos as bullet markers.** Where a bullet focuses on a single named entity (a company, a regulator, a sovereign fund), the entity's logo replaces the bullet marker. Logos auto-fetched from the free [Hunter.io logo API](https://hunter.io/changelog/company-logo-api-free/) and cached locally. No API key, no manual sourcing.
- **Visual system.** Georgia headlines, Calibri body, deep navy `#051C2C`, pure black rules, white background. Charts have data labels at 1 decimal on every bar; gridlines and y-axis numbers are suppressed. No em dashes. No date in footer.
- **Citation discipline.** Every numeric claim gets a source line. Footnotes are distinct from sources and prefixed with `*`. Page numbers in the lower right as `Page X / Y`.
- **15-slide cap** for the main deck. Detail goes to the appendix.

## Use it without Claude

Hand-write a JSON spec following the schema in [`SKILL.md`](SKILL.md) (use `examples/data-center-landscape/input.json` as a template), then run:

```bash
python scripts/build_deck.py path/to/your-spec.json out/your-deck.pptx
```

Pre-warm the logo cache for new companies before rendering:

```bash
python scripts/download_logos.py acme.com partner.com regulator.gov
```

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| Bullets render with `•` instead of a logo | Logo fetch failed (no internet, firewall, or unknown domain) | Run `scripts/download_logos.py <domain>` manually, or check `https://logos.hunter.io/<domain>` returns a PNG in your browser |
| PowerPoint shows "needs to repair" on open | Stale build with float EMU values | Re-render with the latest `build_deck.py` |
| Headline overflows the right edge | Action title exceeds 110 characters at 18pt Georgia | Shorten the headline; the build script prints a warning naming the slide |
| Chart bars have no value labels | Chart family the renderer does not yet style | Bar charts (`content_type: "bar_chart"`) are fully styled; other families fall back to defaults |

## Examples in this repo

- [`examples/data-center-landscape/`](examples/data-center-landscape/): industry report on the data center landscape, 15 slides, rendered both with the skill (`skill-version.pptx`) and with a generic AI-deck-tool baseline (`vanilla.pptx`). This is the deck behind the comparison images above.
- [`examples/market-entry/`](examples/market-entry/): Vietnam joint-venture recommendation, 14 slides including appendix. Demonstrates the recommendation action table.

## Contributing

Pull requests welcome, especially for additional chart families, more worked examples, and theme variations. Please open an issue first for major changes.

## License

MIT. See [LICENSE](LICENSE).
