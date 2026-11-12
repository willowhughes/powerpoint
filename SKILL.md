---
name: powerpoint
description: Create and edit technical presentations as editable PowerPoint files, with custom visuals, consistent formatting and no AI-generated look. Use when the user wants slides, a deck, a talk, a lecture or a seminar on any technical topic, wants to change an existing .pptx, or wants to work in their own slide template, including decks moved through Google Slides.
disable-model-invocation: true
---

# Technical slides

Make decks that look like a professor who knows the subject built them by hand: simple, exact and consistent, with figures designed for each idea. Everything stays editable in PowerPoint.

Look at `examples/contact.png` before your first slide. It shows the quality bar. Don't copy its figures or layouts unless they suit the idea.

## Running the tools

Everything is in this skill's folder, the one holding this file: helpers in `slidekit/`, the default template in `assets/`, the helper reference in `references/helpers.md`, and an example in `examples/`. Work wherever the user's deck is, and run the tools from there:

```
uv run --project <this folder> slidekit <command>
```

The first run sets up its own dependencies. Below, `slidekit` is short for that whole command. If `uv` isn't available, ask the user to install it, or run `pip install -e <this folder>` and call `slidekit` directly.

## Visuals come first

Decide what the audience should see that words can't show, then draw exactly that.

- **Tailor every figure.** It shows how this idea works. If it could illustrate a different idea, it's too generic. Use boxes and arrows only when something really flows.
- **Let shapes carry data.** Draw to scale, use real values, and put the numbers on the figure.
- **Give color a meaning.** Each concept gets one color and keeps it for the whole deck, including in the sentence beside the figure. Gray means inactive. Use one strong color for "look here".
- **Keep the drawing simple.** Pale fill with a darker outline in the same color, thin arrows, white background, and labels on the figure rather than a legend.
- **Reuse and grow figures.** Bring the same figure back as the idea develops, dim what's done and highlight what's new.
- **Lead with the figure,** then one to three bullets on what to notice. Bullets shouldn't repeat what the figure already says.
- **Use equations when they're clearest.** Center them, label parts with underbraces, then define the symbols.

## Words

- **Titles** are short, fit on one line, and name the topic or state the point: "Swing equation: the gap sets the slope".
- **Bullets** are fragments on one line each, with no final period, at most two levels and about five top-level items. Bold the lead term when defining several things.
- **Plain words.** Write what an expert would say out loud.
- **One small italic line** near the bottom for a caveat, what's simplified, or a pointer to later material.
- **Source citations** aren't visible by default; include them when requested or appropriate. Otherwise keep source provenance in speaker notes or beside the build script.

## Truth

- Every number comes from a source, the user, or code you ran. Otherwise mark it illustrative or ask. Never invent data, results or citations.
- Compute plots and worked examples in code rather than typing numbers in.
- Say on the slide what you simplified.
- Tell the user about anything you couldn't verify.

## Deck structure

- Open by orienting the audience with a few questions or a roadmap. Bring it back later with finished parts dimmed.
- Carry one running example through the deck.
- Give the simple version first, say what it leaves out, then add the complications.
- Use a divider only when the topic really changes. End big sections with a small recap table.
- Humor is fine if it's dry and rare.

## Never

These are what make slides look AI-generated.

- Accent lines under titles, color bars, side stripes, gradients, shadows, 3D, icons in circles, clip art, stock photos, emoji
- Card grids, stat tiles, "Key takeaways" boxes, callout boxes around ordinary sentences
- Titles like "Introduction", "Overview", "Conclusion", "The Power of X", "Why X Matters"
- Hype words like leverage, seamless, robust, game-changing, unlock
- Decorative figures, paragraphs on slides, every bullet in bold, centered body text
- The same layout on every slide, or a different look on every slide
- A "Thank you / Questions?" slide, unless asked

## Consistent and editable

- Build on the template's layouts and placeholders, keeping formatting consistent.
- Use theme colors and inherited fonts only, so a template swap restyles everything.
- Keep to a few text sizes: the template's title and bullet sizes, plus about 16 to 20 pt for figure labels and notes.
- Name every meaningful shape and group related ones. That's how people and later agents find things.
- Prefer native objects to pictures, and tell the user whenever something ends up as a picture.

| Content | Use | Editable in PowerPoint? |
|---|---|---|
| Diagrams | Native shapes, with arrows attached to them | Fully |
| Equations | `eq()`: LaTeX becomes a PowerPoint equation | Fully |
| Simple plots | Native charts | Fully, including data |
| Complex plots | Matplotlib image in theme colors, with the script kept beside the deck | Regenerate from the script |
| Tables | `table()`, horizontal rules only | Fully |
| Code | `code()`, plain monospace | Fully |
| Photos, screenshots | Pictures | Move, resize, crop |

## Making a new deck

1. **Ask only what you can't infer:** audience, length, goal, source material, and whether they have a template.
2. **Propose an outline and wait for approval:** slide titles, plus one line per figure on what it should make obvious. This is where the user's judgment goes in, so keep it cheap to change.
3. **Inspect the template:** `slidekit inspect template.pptx`. The default is `slidekit.TEMPLATE`.
4. **Build with a throwaway Python script,** kept beside the deck: `slidekit run build.py`. The script can `import slidekit`. The helpers are in `references/helpers.md`. Use `python-pptx` directly whenever that's simpler; nothing requires the helpers.
5. **Render and look:** `slidekit render deck.pptx`, then open every slide image. Check for overflow, collisions, clipped plots, labels far from what they name, and anything an expert would be embarrassed by. Fix it and render again.
6. **Report plainly:** what's illustrative, what's a picture, and what you couldn't verify.

## Changing an existing deck

The `.pptx` is the source of truth, and people may have edited it by hand. Never rebuild it from an old script.

1. **Back it up:** `slidekit backup deck.pptx`.
2. **See what's there:** `slidekit inspect deck.pptx`.
3. **Change only what was asked.** Leave everything else alone. If you notice another problem, mention it instead of fixing it.
4. **Check your work:** `slidekit compare deck.pptx`, then look at `deck_changes.png`, which shows changed slides old and new. If more changed than was asked, run `slidekit restore deck.pptx` and redo it smaller.

## Working in someone's template

Keep its look; the philosophy above still applies. Layout names and positions vary, so inspect first, and `a1` to `a6` become that template's accent colors. Decoration such as banners and logos isn't listed with the placeholders, so check the render to make sure figures fit around it.

## Google Slides

The user downloads the deck (File > Download > Microsoft PowerPoint), you edit that file, and they bring back only the changed slides (File > Import slides). Tell them to check equations and charts after importing, since they may arrive simplified or as static images. For decks that live in Google Slides, write simple math as styled text and insert harder equations as pictures made with matplotlib.
