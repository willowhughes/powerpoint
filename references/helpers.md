# Helpers and PowerPoint gotchas

`import slidekit` gives thin helpers over `python-pptx`. Use them when they save time and ignore them when they don't. `examples/grid_frequency.py` shows them in use.

Positions and sizes are in inches. `parent` is a slide or a group, so `g = slide.shapes.add_group_shape()` then `box(g, ...)` builds a group in slide coordinates.

**Colors** are theme slots: `ink`, `paper`, `a1` to `a6`. Add `-` for darker, `+` for a pale tint, `~` for muted. On the default template, `a1` is blue, `a2` orange, `a3` green, `a4` pink, `a5` purple, `a6` amber.

**Text markup:** `**bold**`, `*italic*`, `_{sub}`, `^{sup}`, `[a3:colored]`. Write simple math in the body font like `*x_{t}*`, and save `eq()` for real equations.

| Helper | What it does |
|---|---|
| `TEMPLATE` | Path to the default template |
| `open_deck(path)` | Opens a `.pptx` or `.potx`, which `python-pptx` alone refuses |
| `add_slide(prs, title, layout, footer)` | New slide, keeping the footer and slide number that `python-pptx` drops |
| `bullets(slide, items, x, y, w, h, anchor)` | Fills the body placeholder. Items are `(level, text)` or `(level, text, points_before)` |
| `text(parent, markup, x, y, w, h, size, color, align, anchor)` | Text box with no padding, so it lines up with shapes |
| `note(slide, markup, size, lines)` | Small line just above the footer, always in the same place |
| `code(parent, source, x, y, w, h, size)` | Monospaced code, exactly as written |
| `box(parent, x, y, w, h, color, label)` | Rounded box with a pale fill and darker outline |
| `rect(parent, x, y, w, h, fill, line, shape=...)` | Rectangle or preset shape, for marks whose size carries data |
| `outline(parent, x, y, w, h)` | Unfilled box for marking a region |
| `circle(parent, cx, cy, r, fill, line, label)` | Circle. `shape=MSO_SHAPE.FLOWCHART_OR` draws a circled plus |
| `arrow(parent, a, b, color, width, head)` | Connector. Ends are `(x, y)`, or `(shape, "top")` to stay attached when the shape moves |
| `path(parent, steps, color, width, fill, head)` | Custom curve: `("M", x, y)`, `("L", x, y)`, `("C", x1, y1, x2, y2, x, y)`, `("Z",)` |
| `eq(slide, latex, x, y, w, h, size, align, display, fallback)` | LaTeX to a native equation. `display=False` is compact and inline |
| `equations(slide)`, `replace_eq(slide, i, latex)` | Find or swap equations, which `slide.shapes` can't see |
| `table(slide, rows, x, y, widths, row_h, size, align)` | Native table with horizontal rules only |
| `style_chart(chart, font, legend, title)` | Removes the default chart look |
| `axis_title(axis, markup)` | Axis title in the same style |
| `fix_plot(frame, x, y, w, h)` | Pins the plot area to slide coordinates, so overlays can be placed from data |
| `theme_colors(prs)` | Theme colors as hex, for matplotlib |
| `duplicate_slide(prs, slide)`, `delete_slide(prs, slide)`, `move_slide(prs, slide, i)` | What `python-pptx` leaves out |

## Gotchas

**Layout**
- Read `prs.slide_width` rather than assuming a size. The default template is 13.33 x 7.5 inches.
- A negative width or height makes PowerPoint report a corrupt file. `slidekit render` catches this first.
- Templates put footers and slide numbers in unexpected places, sometimes at the top.
- Use the "Title Only" layout when there are no bullets, because empty placeholders show prompts while editing.
- Many templates ship with sample slides. Remove them with `delete_slide` after building.

**Text**
- `text_frame.text = ...` wipes formatting. Change `run.text` to keep it.
- `*` and `_{` are markup, so anything full of them needs `code()` or `text(..., raw=True)`.

**Equations**
- Check every equation in the render. Fractions can make rows taller than you expect.
- `slide.shapes` can't see them, so `slidekit inspect` lists them separately.
- `fallback` is what apps without native equations show. Give it a readable version.

**Charts**
- Build with `CategoryChartData` or `XyChartData`, then call `style_chart`.
- Color single bars with `series.points[i].format.fill` and lines with `series.format.line.color`.
- For overlays such as markers or a box around some bars, call `fix_plot` and compute positions from the data. Leave at least 0.2 inches between the plot and the chart's frame, or PowerPoint shifts the plot.
- Overlays are separate shapes, so they don't follow if someone moves the chart. Tell the user.

**Pictures**
- `add_picture` doesn't take SVG. Use PNG at 300 dpi.
- For matplotlib, take colors from `theme_colors(prs)`, save with `transparent=True`, and set the picture's alt text through `pic._element.nvPicPr.cNvPr.set("descr", ...)`.

**Step-by-step builds**
- `python-pptx` can't write animations. Make a series of slides instead: draw the same figure with a step argument, or use `duplicate_slide` and change the copy. Remember that a later edit then has to be repeated on each one.

**Speaker notes**
- `slide.notes_slide.notes_text_frame.text = "..."`.

**Rendering**
- PowerPoint is used on Windows when installed; otherwise LibreOffice. LibreOffice is an approximation, so tell the user to check fonts and equations in PowerPoint before presenting.
