# AI Agents: a simple guide

A talk for the HR team: what an AI agent is, a real eight-agent team running every day, an app built with an agent, and a proposal for an AI team for HR.

It comes in two versions with the same content: the web deck is the main one, and the PowerPoint is its twin.

| | File | Motion |
|---|---|---|
| **Web deck** | `index.html` | Every slide has its own entrance: headlines rise from behind a mask, the agent loop draws itself, the 8 rolls up like an odometer, and connector lines trace from CEO to the team and from Head of HR to the specialists |
| **PowerPoint** | `pptx/AI-Agents.pptx` | A fade transition on every slide, then each element fades in on its own timing, with no clicks needed |

## Presenting the web deck

Open `index.html` in Chrome, Edge, Safari or Firefox. It runs offline from the folder, with no server and no internet.

| Key | Action |
|---|---|
| `→` `Space` `Enter` | Next slide |
| `←` `Backspace` | Previous slide |
| `F` | Full screen |
| `N` | Speaker notes overlay |
| `P` | Presenter window (notes, timer, next slide), synced with the main window |
| `1`–`9` `Home` `End` | Jump |
| `?` | Shortcut help |

You can also click the right side of the slide to go forward and the left side to go back, or swipe on a phone or tablet. Click any screenshot to zoom it full screen.

**Save as PDF:** press Ctrl/Cmd + P and choose *Save as PDF*. You get one page per slide, with every animation in its final state.

## The design

- **Black, white and gray only.** The screenshots keep their real colors because they're evidence.
- **Type:** Schibsted Grotesk for the voice, with one EB Garamond italic phrase per headline as the accent. The PowerPoint uses Arial and Georgia so it looks the same on any computer.
- The light content slides are bookended by dark slides: the cover, the "8 agents" moment, Dojima and the close.
- Motion uses exponential ease-out, and nothing moves just to move. If the viewer's system asks for reduced motion, the deck switches to simple fades.

## Rebuilding the PowerPoint

The PowerPoint is generated from `pptx/build.cjs`. After you edit copy there:

```bash
npm install
npm run pptx
```

## Files

```
index.html            the web deck: all 14 slides and their speaker notes
assets/deck.css       design system, layouts and the entrance choreography
assets/deck.js        navigation, presenter sync, lightbox, drawn connectors
assets/icons.js       Lucide icons (ISC licence)
assets/fonts/         self-hosted Schibsted Grotesk + EB Garamond (OFL)
assets/img/           the screenshots
pptx/                 PowerPoint build script and output
.claude/              design skills and agents for Claude Code
```
