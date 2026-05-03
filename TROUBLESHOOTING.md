# Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| `ImportError: No module named 'pptx'` | python-pptx not installed | `pip install python-pptx` (or `python3 -m pip install python-pptx`) |
| Bullets render with `•` instead of a logo | Logo fetch failed (no internet, firewall, or unknown domain) | Run `scripts/download_logos.py <domain>` manually, or check `https://logos.hunter.io/<domain>` returns a PNG in your browser |
| PowerPoint shows "needs to repair" on open | Stale build with float EMU values from an older script version | Re-render with the latest `build_deck.py` |
| Headline overflows the right edge | Action title exceeds 110 characters at 18pt Georgia | Shorten the headline; the build script prints a warning naming the slide |
| Chart bars have no value labels | Using a chart family the renderer does not yet style | Bar charts (`content_type: "bar_chart"`) are fully styled; other families fall back to defaults |
| Source URLs are not clickable | Source string did not use `[text](url)` markdown syntax | See the citation format section in [`SKILL.md`](SKILL.md) |
| Source link silently missing from the rendered deck | URL returned HTTP 404 or 410 at render time | The build script hides definitively-broken links by design. Replace the URL with one that resolves, or remove `(url)` to render as plain text. |

## When to file an issue

If you hit something not in this table, open an issue at <https://github.com/floflo11/mbb-decks/issues> with your `input.json` and the full error output.
