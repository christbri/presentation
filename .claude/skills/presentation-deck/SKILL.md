---
name: presentation-deck
description: House style and process for building presentation decks for this user. Use whenever making, redesigning or revising a slide deck or presentation (.pptx), including turning a draft deck into a polished one.
---

# Presentation decks: what this user wants

Learned from building `pptx/AI-Agents.pptx` (this repo). That deck is the reference: read `pptx/build.cjs` before starting a new one and reuse its helpers.

## Format
- **PowerPoint (.pptx) only.** Don't build HTML/web decks; the user rejected one ("sampah").
- Generate with pptxgenjs from a build script (`pptx/build.cjs` pattern), on a 1920×1080 grid (144 px = 1 in, 2 px = 1 pt), `LAYOUT_WIDE`.
- Safe fonts so it renders the same everywhere: Arial for the voice, Georgia italic for one accent phrase per headline.

## Look
- Minimalist and aesthetic. Black, white and gray only unless told otherwise; screenshots keep their real color.
- No eyebrow/kicker labels above titles and no "01 ·" section numbers; a quiet footer (section name + `NN / total`) and a thin progress line instead.
- Dark slides open, close, and mark big moments; light slides carry content.
- Never leave big empty areas next to a title on the cover or closing slide. Fill them with **original art**, e.g. generative halftone pieces (`pptx/art/make_art.py`: dot sphere with wired nodes, dot-built question mark).
- Never repeat or copy a diagram from another slide as decoration. Each visual must be distinct ("jgn copas").

## Motion
- **Morph on every slide**, not fades or per-element entrance animations. The user explicitly wants connected motion where each slide flows into the next.
- Technique: one persistent black shape named `!!ink` (always `roundRect`; vary size and corner radius, a square with max radius = circle) that becomes a different element on each slide: full-bleed dark background, a circle, a panel, a rule, a box, a speech bubble, a frame. Also name titles `!!title`/`!!title2`, the counter `!!count`, the progress bar `!!progress`.
- Inject the transition into each slide XML after `pres.write()`: `mc:AlternateContent` with `p159:morph option="byObject"`, `p14:dur` ≈ 1400, fallback `<p:fade/>`. The first slide gets a plain fade.

## Content
- Keep every piece of the user's material: all copy, every screenshot, every speaker note (`slide.addNotes`).
- Don't put personal names on slides (e.g. agent names); roles only.
- No "press to begin" / "click to view" hints.

## Process
- Ask once up front about anything genuinely ambiguous (palette words, typos), then build.
- QA every slide: `validate.py`, render via LibreOffice (install `libreoffice-impress` if missing), inspect, fix, re-render. The preview substitutes Georgia with a wider font, so leave ~10% slack and tell the user what to double-check in real PowerPoint (Morph can't be previewed here).
- Commit and push to the session branch, then send the .pptx with SendUserFile.
