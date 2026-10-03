# AI Agents: a simple guide

A talk for the HR team: what an AI agent is, a real eight-agent team running every day, an app built with an agent, and a proposal for an AI team for HR.

**The deck: [`pptx/AI-Agents.pptx`](pptx/AI-Agents.pptx)**: 14 slides, black, white and gray, with speaker notes on every slide.

## The motion: Morph, end to end

Every slide uses PowerPoint's **Morph** transition, so each slide grows out of the one before it.

One black shape travels through the whole deck and turns into something new on every slide:

| Slide | The black shape becomes |
|---|---|
| 1 · AI Agents | the whole dark cover |
| 2 · An AI that gets things done | the **Your goal** circle at the centre of the loop |
| 3 · Chat AI vs AI Agent | the **AI Agent** panel |
| 4 · What an agent can do | the rule above the six capabilities |
| 5 · What every agent is made of | the **Model** bar under Soul, Skills and Memory |
| 6 · My own AI team | the whole dark slide behind the 8 |
| 7 · How my team is structured | the **CEO** box |
| 8 · Their office, live | the frame around the office screenshot |
| 9 · One request, a whole team | the request speech bubble |
| 10 · Real output, every morning | the panel holding both Discord screenshots |
| 11 · From my notes to a clean journal | the arrow between notes and journal |
| 12 · Dojima | the whole dark slide |
| 13 · An AI team for HR | the **Head of HR** box |
| 14 · Start small | the whole dark closing slide |

The titles, the slide counter and the progress line along the bottom morph between slides as well.

Morph needs PowerPoint 2019, PowerPoint for Microsoft 365, or PowerPoint for the web or mobile. Older versions fall back to a fade. Keynote and Google Slides don't play Morph.

## Editing

You can edit text directly in PowerPoint. If you add your own shapes and want them to morph too, give them the same name on both slides, starting with `!!` (Home → Arrange → Selection Pane). The travelling black shape is called `!!ink`.

To regenerate the file from code:

```bash
npm install
npm run pptx
```

## Files

```
pptx/AI-Agents.pptx   the deck
pptx/build.cjs        builds the deck (layout, copy, notes, Morph)
assets/img/           the screenshots
assets/icons.js       Lucide icons (ISC licence)
.claude/              design skills and agents for Claude Code
```
