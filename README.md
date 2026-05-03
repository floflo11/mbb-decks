# MBB Decks

> Real consulting decks from Claude. Not slop.

Built by [Iris Meng](https://www.linkedin.com/in/yilin-meng/), co-founder of New York AI Labs, working at the intersection of finance and AI.

## Without skill vs with skill

Same prompt, same topic (a data center industry report), same author. **Left** is what Claude produces by default. **Right** is what Claude produces with this skill installed.

![Cover comparison](examples/data-center-landscape/compare-cover.png)

![Agenda comparison](examples/data-center-landscape/compare-agenda.png)

![Content comparison](examples/data-center-landscape/compare-content.png)

![Closing comparison](examples/data-center-landscape/compare-closing.png)

Both decks are committed to the repo. Open them in PowerPoint or Keynote and flip through.

## Install

In [Claude Code](https://claude.com/claude-code):

```
/plugin marketplace add floflo11/mbb-decks
/plugin install mbb-decks
```

Then prompt:

> Build me an MBB-style deck on [your topic]. Use the mbb-decks skill.

The first time the skill runs it self-checks for Python and `python-pptx`, and tells you the one command to run if anything is missing.

## Star to follow

<p align="center">
  <img src="docs/star-this-repo.png" alt="Click the Star button at the top of this page" width="280" />
</p>

If the comparison above made you smile, **star this repo to follow more Microsoft Office skill releases**. Coming next:

- **Excel** — three-statement models, LBO, sensitivity tables, valuation builds
- **Word** — IC memos, board prereads, one-pagers
- **Outlook** — partner-style email and follow-ups

Stars vote on which ships first.

## What's different

A consulting deck is not slides with bullets. It is an argument.

This skill teaches Claude four house-style conventions:

- **Action titles tell the story.** Read the headlines top to bottom; that is the deck.
- **One claim, one chart, one slide.** With the takeaways panel beside it. Never split a chart from its commentary.
- **Logos for companies, not letters in circles.** Real brand marks, auto-fetched, where they belong.
- **Fifteen slides, hard cap.** Partners do not read further. Detail goes in the appendix.

The exact rules and the JSON schema live in [`SKILL.md`](SKILL.md).

## Examples

- [`examples/data-center-landscape/`](examples/data-center-landscape/) — the deck behind the comparison images above. Industry report with logos on every bullet that names a company.
- [`examples/market-entry/`](examples/market-entry/) — Vietnam JV recommendation. Owner / Action / Outcome table at the close.

## For developers

Skip Claude Code. Render from a JSON spec directly:

```bash
git clone https://github.com/floflo11/mbb-decks
cd mbb-decks && pip install python-pptx
python scripts/build_deck.py path/to/spec.json out.pptx
```

Stuck? See [`TROUBLESHOOTING.md`](TROUBLESHOOTING.md).

## License

MIT. See [LICENSE](LICENSE).
