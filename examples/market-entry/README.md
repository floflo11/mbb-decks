# Example: Vietnam market entry recommendation

A 14-slide board pre-read recommending entry into Vietnam via a joint venture, including five-year P&L projection and partner due diligence in the appendix.

## What this example demonstrates

- Cover slide with title, subtitle, author, organization
- Executive summary with a single full-sentence headline and three MECE bullets
- Three-item agenda with subtitles
- Three section dividers (`01`, `02`, `03`)
- Bar chart slide (clustered column with three series)
- Bullet-list slide
- Two-column comparison slide
- Recommendation slide with a four-row action table (Owner / Action / Outcome)
- Appendix divider plus two appendix slides (financial projection chart, due diligence summary bullets)

## Render it

```bash
python ../../scripts/build_deck.py input.json output.pptx
open output.pptx
```

## Storyline

Read these action titles top to bottom and confirm they form one argument before expanding any body content:

1. Southeast Asian middle-class spending grows from $1.5T to $2.8T by 2030, with Vietnam and Indonesia driving most of the increase
2. Acme's product portfolio aligns with Vietnam's three highest-growth consumer categories, none of which has a dominant Western player today
3. Joint venture is preferred over greenfield and acquisition: lower capital risk, faster regulatory clearance, retained operational control
4. Local Partner X scores highest across distribution, regulatory standing, and category fit, with one open issue on shared capex governance
5. Approve $40M for the Vietnam joint venture with Local Partner X, targeting Q3 2026 launch

This is the ghost deck. The body content (bullets, charts, footnotes, sources) only exists to prove these claims.
