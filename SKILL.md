---
name: mbb-decks
description: Use when generating consulting-grade slide decks, MBB-style presentations, McKinsey/Bain/BCG-format slides, board prereads, executive readouts, recommendation decks, market entry decks, due diligence decks, or any business presentation that requires action-title storylines, pyramid-principle structure, MECE bullets, and a clean editorial visual system rendered to a real .pptx file.
---

# MBB Decks

Generate consulting-grade .pptx decks from a JSON spec. Encodes the MBB house-style rules: action-title storyline, six-section structure, MECE bullets, citation discipline, and a navy-and-white editorial visual system.

## When to use

- User asks for an MBB deck, McKinsey-style deck, consulting deck, board pre-read, executive readout, market entry deck, due diligence deck, recommendation deck, or strategy presentation.
- User provides analysis that needs to be turned into slides for a non-technical executive audience.
- User wants a real .pptx file, not a Markdown outline.

## Workflow: ghost deck first, expand second

Do NOT skip the ghost-deck step. It is the single biggest quality lever.

0. **Verify renderer dependencies (first run only).** Before producing any deck, confirm Python 3.9+ and `python-pptx` are installed by running:
   ```bash
   python3 -c "import pptx; print(f'python-pptx {pptx.__version__} on Python {__import__(\"sys\").version.split()[0]}')"
   ```
   If this prints a version, skip to step 1. If it errors with `ModuleNotFoundError: No module named 'pptx'`, tell the user:
   > The renderer needs `python-pptx`. Run `pip install python-pptx` in your terminal, then say "ready" and I'll continue.
   Wait for the user to confirm before proceeding. The build script also fails gracefully with the same install message if you skip this check, but checking up front gives a cleaner experience.
1. **Ghost deck.** Produce only the action titles for every slide (cover, executive summary headline, agenda items, every content slide's action title, recommendation headline). Read them top to bottom as a paragraph. They must form one coherent argument: situation → complication → resolution, or claim → evidence → recommendation.
2. **Confirm with the user.** Show the ghost deck. Ask: "Does this storyline land? Anything to add, cut, or reorder before I expand?" Wait for confirmation.
3. **Expand.** Fill in bullets, charts, footnotes, sources. Keep each slide MECE under its action title.
4. **Render.** Write the JSON spec to disk and call `scripts/build_deck.py`.

## Deck structure (six sections)

Hard cap: 15 slides for sections 1 to 5. Appendix is unbounded.

| # | Section | Notes |
|---|---------|-------|
| 1 | Cover | Title, subtitle, author, organization. Blank theme. |
| 2 | Executive summary | One headline that answers "so what." Three to five MECE bullets. |
| 3 | Agenda | Two to four sections. Section titles match the dividers. |
| 4 | Sections | Each begins with a divider slide (number + title). Two to four content slides per section. |
| 5 | Recommendation | Headline + action table (Owner / Action / Outcome). Action table is mandatory. |
| 6 | Appendix | Optional. Backup detail and supporting analysis. Same content slide format. |

## Action title rules

Every slide except cover, dividers, and appendix divider has an action title. The action title carries the meaning. The body proves it.

- **One full sentence**, 10 to 15 words. Subject + verb + claim.
- **Specific.** "Vietnam middle-class spending grows 7% CAGR through 2030" beats "Market is attractive."
- **Stand-alone.** A reader skimming only the action titles must understand the argument.
- **Causal where possible.** "X drives Y" beats "X and Y."
- **No questions, no labels.** "Market sizing" is a label, not an action title. Replace with the finding.

Read all action titles in sequence before writing any body content. If the sequence does not tell a coherent story, fix the storyline before expanding.

## MECE bullets

Each slide has 3 to 5 bullets max. They must be:

- **Mutually exclusive.** No overlap.
- **Collectively exhaustive.** No obvious gap under the action title.
- **Parallel structure.** Same grammatical pattern across bullets on a slide.
- **Concrete.** Numbers, dates, named entities. "Three local partners screened" beats "Several partners considered."

## Visual system

Defaults are encoded in `scripts/build_deck.py`. Do not override unless the user explicitly asks.

| Element | Specification |
|---------|---------------|
| Background | White `#FFFFFF` |
| Headlines | Georgia, deep navy `#051C2C` |
| Body text | Calibri, near-black `#1A1A1A` |
| Accent | Electric blue `#2251FF`, used sparingly for rules and section numbers |
| Gridlines, dividers | Light grey `#E5E7EB` |
| Footnotes, sources, page numbers | Calibri 8pt, mid-grey `#949BA8` |
| Slide size | 16:9, 13.333" × 7.5" |
| Margins | 0.5" left and right |
| Page number | Lower right, format `Page X / Y` |
| Date in footer | Never |
| Icons | Never, unless user explicitly requests |

## Chart families

Pick the chart that matches the claim in the action title.

| Action title verb | Chart family |
|-------------------|--------------|
| "Grows," "declines," "trended" | Column or line |
| "Composed of," "share of" | 100% stacked column or pie (pie used sparingly) |
| "Higher than," "ranks first" | Bar (horizontal) sorted by value |
| "Drives," "explains" | Waterfall or stacked bar |
| "Correlates with" | Scatter |
| "Builds from X to Y" | Waterfall |

The script currently renders column-clustered charts cleanly. For other chart types, fall back to bullets and describe the data in prose, then add the chart in PowerPoint manually OR extend the script.

## Citation format

- **Inline marker** in body text: `[1]`, `[2]`, etc., when a specific number or claim needs attribution.
- **Source line** at the bottom of the slide: `Source: [text](url); [text](url)`. Use markdown-style `[text](url)` syntax so each source is rendered as a clickable hyperlink in the `.pptx`. Multiple sources separated by semicolons. Plain text without `(url)` still works for sources that have no public URL.
- **Footnote line** above the source: explanatory note about a number (e.g., "Excludes Singapore"). Distinct from source.
- No source line if the slide makes no specific data claim.

## Data sourcing rules (critical)

The skill is only as credible as its numbers. These rules apply to every chart, every bullet, every claim with a number in it.

### Always use the current year

Today's date is available in your runtime context. Use it.

- For historical data: cite the most recent available release. If a 2026 dataset exists, do not cite the 2023 version.
- For forward-looking projections: forecast labels should start from the current year forward. If today is in 2026, forward years are `2026E`, `2027E`, `2028E`. Never label a past year as `E` (estimate).
- For "as of" qualifiers in footnotes: say "as of Q[current quarter] [current year]", not a stale date.

### Use only authoritative, recognized sources

A short non-exhaustive list of sources Claude should reach for, by domain:

| Domain | Authoritative sources |
|--------|----------------------|
| Macro and trade | World Bank, IMF, OECD, BIS, UN ComTrade, WTO |
| Cloud and AI infrastructure | Synergy Research Group, IDC, Dell'Oro Group, Gartner, JLL Data Center Outlook, Structure Research |
| Energy and climate | IEA, BloombergNEF, EIA, Lazard LCOE, NREL ATB, Wood Mackenzie |
| Financial markets | Bloomberg, S&P Global, Refinitiv (LSEG), FactSet, Capital IQ, PitchBook, Preqin |
| Consulting/strategy reports | McKinsey Global Institute, BCG, Bain (publicly cited), Oliver Wyman |
| Consumer | Euromonitor, Nielsen, Kantar, eMarketer, Statista |
| Public companies | Company 10-K, 10-Q, S-1, earnings transcripts (SEC EDGAR) |
| Regulatory and policy | Federal Reserve, BLS, BEA, ECB, EU Commission, regulator filings |

Do NOT use: random blog posts, AI-generated summaries, paywalled content without a verifiable cite, or "Statista projections" as the only source for a numeric claim.

### Every source needs a clickable URL when one exists

Format: `Source: [BloombergNEF Energy Transition Investment Trends, 2025](https://about.bnef.com/energy-transition-investment/); [IEA World Energy Outlook, 2024](https://www.iea.org/reports/world-energy-outlook-2024)`

The build script automatically:
- Renders each `[text](url)` segment as a hyperlink in the `.pptx`.
- Sends a HEAD request to each URL at render time. If the URL returns a 4xx error, the link is dropped silently and only the plain text is rendered. The reader does not see a broken link.
- Leaves the link in place for non-4xx outcomes (network timeout, DNS error, blocked HEAD), since we cannot definitively call those broken.

If you cannot find a public URL for a specific report, leave the source as plain text. Do not invent URLs.

## Pairing chart with commentary

Never ship a chart slide as just a chart. Every chart slide should carry its own commentary panel on the same slide so the reader gets the picture and the takeaway in one view. If the bullets describing a chart end up on the next slide, merge them.

Two layouts:

- **`commentary_position: "right"`** (default). Chart takes left 62%, "KEY TAKEAWAYS" panel on the right with three to four bullets. Use for most chart slides.
- **`commentary_position: "below"`**. Chart full-width on top (~2.8" tall), commentary in a 2-column row below. Use when the chart needs more horizontal room (5+ categories with multiple series), or when there are 4+ takeaway bullets.

## Logos as bullet markers

Wherever a commentary or two-column bullet focuses on a single named entity (company, regulator, sovereign fund), use the entity's logo as the bullet marker instead of an abstract icon or a plain dot. The skill auto-downloads logos from logo.dev (a public token ships with the skill; override with `LOGODEV_TOKEN` for a private quota) and caches them in `assets/logos/`. If logo.dev does not yet have a brand mark for the domain it returns HTTP 202 and the bullet falls back to a plain `•` marker rather than a placeholder.

A bullet entry can be:

- A plain string → renders with `•` marker.
- `{ "icon": "→", "text": "..." }` → renders the glyph in navy as the marker.
- `{ "logo": "microsoft.com", "text": "..." }` → renders the company logo as the marker. Auto-downloaded on first render and cached.
- `{ "logo": "...", "icon": "...", "text": "..." }` → logo wins.

If a bullet covers a concept rather than an entity (a region, a principle), use `icon`, not `logo`. If it names multiple companies, split the bullet so each has one anchor entity.

## JSON spec

The script reads a JSON file with this shape. See `examples/market-entry/input.json` and `examples/data-center-landscape/input.json` for complete worked examples.

```jsonc
{
  "meta": { "title": "...", "subtitle": "...", "author": "...", "organization": "..." },
  "executive_summary": {
    "headline": "Full-sentence claim, 15 to 25 words",
    "bullets": ["...", "...", "..."]
  },
  "agenda": [
    { "title": "Section title", "subtitle": "Optional subtitle" }
  ],
  "sections": [
    {
      "title": "Matches an agenda title",
      "slides": [
        {
          "action_title": "Full sentence, 10 to 15 words",
          "content_type": "bullets",       // or "bar_chart" or "two_column"

          // for content_type "bullets"
          "bullets": [
            "Plain bullet text",
            { "icon": "$", "text": "Bullet with abstract icon marker" },
            { "logo": "microsoft.com", "text": "Bullet with company logo marker" }
          ],

          // for content_type "bar_chart"
          "chart": {
            "categories": ["A", "B", "C"],
            "series": [{ "name": "...", "values": [1.0, 2.0, 3.0] }],
            "unit": "USD millions"
          },
          "commentary": [                  // optional; pairs with chart
            { "logo": "aws.amazon.com", "text": "AWS-specific takeaway" },
            { "logo": "microsoft.com", "text": "Azure-specific takeaway" },
            { "icon": "◆", "text": "General takeaway not tied to one company" }
          ],
          "commentary_position": "right",  // or "below"; default "right"

          // for content_type "two_column"
          "left_header": "...",
          "left":  [ "plain", { "icon": "$", "text": "..." }, { "logo": "...", "text": "..." } ],
          "right_header": "...",
          "right": [ /* same shape as left */ ],

          "footnote": "Optional. Auto-prefixed with '*'. Body or chart label should carry a matching '*' marker.",
          "source": "Optional. Format: '[Org], [Publication], [Year]; [Org 2], [Year]'"
        }
      ]
    }
  ],
  "recommendations": {
    "headline": "Approve / proceed / invest claim",
    "actions": [
      { "owner": "CFO", "action": "Specific action by date", "outcome": "What it produces" }
    ]
  },
  "appendix": [
    // Same shape as section slides
  ]
}
```

## Calling the build script

```bash
python scripts/build_deck.py <input.json> <output.pptx>
```

The script validates the slide count and prints a warning if the main deck exceeds 15 slides.

## Common mistakes

| Mistake | Fix |
|---------|-----|
| Action title is a label ("Market sizing") | Rewrite as a sentence with a verb and claim |
| Action titles do not connect into a story | Reorder or rewrite. Read them all aloud as one paragraph. |
| More than five bullets per slide | Cut to three or split the slide |
| Bullets repeat the action title | Bullets must add evidence, not restate the claim |
| Chart slide with no commentary | Add a `commentary` panel on the same slide; never let bullets describing a chart spill onto the next slide |
| Chart is decorative, not load-bearing | Replace with the chart that proves the action title |
| No source on a data slide | Add one. Every numeric claim needs attribution. |
| Decorative icons sprinkled into body text | Remove. Use icons or logos only as bullet markers, never as page decoration |
| Bullet names a company but uses an abstract icon | Replace with `{"logo": "<domain>", "text": "..."}`; the logo becomes the marker |
| Mixing logo markers and navy-circle icons in the same panel | Drop the circle background on icon glyphs (the script does this automatically) |
| Em dashes | Replace with periods or commas. House style is no em dashes. |
| Date in footer | Remove. House style omits date. |
| Slide count exceeds 15 in main deck | Move detail to appendix. |

## Red flags that mean stop and restart

- "I'll write the slides first and figure out the story after." Stop. Ghost deck first.
- "The action title is too long, I'll shorten it." Check it is still a complete claim. If a label is shorter, that is not the right tradeoff.
- "Three bullets is not enough, I have more to say." Split into two slides or move detail to appendix.
- "I'll put the chart on one slide and the bullets describing it on the next." Stop. Merge them. Add a `commentary` panel on the same slide.
- "I'll add a few icons to make it visual." Use logos for bullets that name a company. Use abstract icons only for concept bullets. Never sprinkle decorative icons.
