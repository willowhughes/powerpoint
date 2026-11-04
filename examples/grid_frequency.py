"""Example deck: why grid frequency falls. Run with: slidekit run grid_frequency.py

A reference for quality, not a template. Every number on the slides comes from the model below.
"""
import inspect
import logging
import math
import sys
from pathlib import Path

import matplotlib
import numpy as np
from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt

from slidekit import (TEMPLATE, I, add_slide, arrow, axis_title, box, bullets, circle, code, eq, fix_plot, note,
                      paint, rect, runs, style_chart, table, text, theme_colors)

matplotlib.use("Agg")
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
import matplotlib.pyplot as plt  # noqa: E402

FOOTER = "Grid frequency"
HERE = Path(__file__).parent
LOSS, F0, LAG = 0.03, 60, 4
REFERENCE = 59.5


def simulate(H, T, loss=0.03, R=0.2, D=1.0, f0=60.0, dt=0.01, t_end=60.0):
    df, pg = 0.0, 0.0                              # deviation, governor power (pu)
    out = []
    for t in np.arange(0, t_end, dt):
        pg += dt * (-df / R - pg) / T              # governor: droop with a lag
        df += dt * (pg - loss - D * df) / (2 * H)  # swing equation with load damping
        out.append((t, f0 * (1 + df)))
    return np.array(out)


def nadir(run):
    i = run[:, 1].argmin()
    return run[i, 0], run[i, 1]


def slope(H):
    return F0 * LOSS / (2 * H)


HIGH, LOW = simulate(5, LAG), simulate(2, LAG)
T_HIGH, F_HIGH = nadir(HIGH)
T_LOW, F_LOW = nadir(LOW)
SETTLE = HIGH[-1, 1]


def questions(prs, dim=0):
    s = add_slide(prs, "Four questions", footer=FOOTER)
    items = [("What sets frequency?", "The balance of power in and out"),
             ("How fast does it fall?", "The gap divided by inertia"),
             ("What stops the fall?", "Governors, then slower reserves"),
             ("What changes with less inertia?", "Faster fall, deeper low")]
    bullets(s, [(0, f"[ink~:**{q}** {a}]" if i < dim else f"**{q}** {a}") for i, (q, a) in enumerate(items)])
    note(s, f"*Running example: a {F0} Hz system loses {LOSS:.0%} of its generation at t = 0*")


def balance_panel(s, x, tripped):
    """Power flowing through the spinning machines, with band widths to scale."""
    thick, top, cx, r, end = 1.3, 2.05, x + 2.6, 0.62, x + 5.3
    gen = 1 - LOSS if tripped else 1.0
    text(s, "**A plant trips**" if tripped else "**In balance**", x, 1.25, 5.3, 0.36, size=22, name="Panel title")
    text(s, f"Generation {gen:.0%}", x, 1.68, 1.75, 0.32, size=16, name="Generation label")
    text(s, "spinning machines", cx - 0.9, 1.68, 1.8, 0.32, size=16, align="c", name="Machines label")
    text(s, "Load 100%", cx + 1.1, 1.68, 1.6, 0.32, size=16, name="Load label")
    flow = rect(s, x, top, end - x, thick * gen, fill="a1+", line="a1", shape=MSO_SHAPE.PENTAGON,
                name="Power from generation (to scale)")
    flow.adjustments[0] = 0.28 / (thick * gen)
    if tripped:
        lost = rect(s, x, top + thick * gen, cx - x, thick * LOSS, fill=None, line="ink", name="Tripped plant")
        lost.line.dash_style = MSO_LINE.DASH
        text(s, f"lost plant: {LOSS:.0%}", x, top + thick + 0.06, 2.0, 0.3, size=16, name="Lost label")
        rect(s, cx, top + thick * gen, end - cx - 0.28, thick * LOSS, fill="a2", name="Kinetic energy (to scale)")
        text(s, f"[a2-:**{LOSS:.0%} from kinetic energy**]", end - 3.2, top + thick + 0.06, 2.92, 0.3, size=16,
             align="r", name="Kinetic energy label")
    circle(s, cx, top + thick / 2, r, line="ink", width=1.25, label="[a2-:**falling**]" if tripped else "60 Hz",
           size=18, name="Spinning machines")


def balance(prs):
    s = add_slide(prs, "Frequency measures the power balance", footer=FOOTER)
    balance_panel(s, 0.75, False)
    balance_panel(s, 7.05, True)
    bullets(s, [(0, "Frequency is the speed of the synchronized machines"),
                (0, "Any gap is paid from their kinetic energy, so they slow")],
            0.75, 4.35, 11.85, 1.5, anchor="t")
    note(s, f"*Band widths are to scale: a {LOSS:.0%} gap looks small, and it is enough*")


def swing(prs):
    s = add_slide(prs, "Swing equation: the gap sets the slope", footer=FOOTER)
    eq(s, r"\underbrace{\frac{2H}{f_0}}_{\text{inertia}}\;\underbrace{\frac{df}{dt}}_{\text{slope}}"
          r"\;=\;\underbrace{P_{\text{gen}}-P_{\text{load}}}_{\text{the gap}}",
       0.75, 1.15, 11.85, 1.5, size=32, anchor="t", fallback="(2H/f0) df/dt = Pgen - Pload")
    bullets(s, [(0, "*H*: stored kinetic energy per unit of capacity, in seconds"),
                (0, "Governors have not moved yet, so at first the gap acts alone")],
            0.75, 2.85, 11.85, 1.3, anchor="t")
    eq(s, r"\left.\frac{df}{dt}\right|_{t=0}=\frac{f_0\,\Delta P}{2H}=\frac{%d\times(-%.2f)}{2\times 5}=-%.2f\ \text{Hz/s}"
       % (F0, LOSS, slope(5)),
       0.75, 4.3, 11.85, 1.2, size=28, fallback=f"df/dt at t=0 = f0 dP / 2H = -{slope(5):.2f} Hz/s")
    note(s, "*One-machine equivalent of the whole system; powers are per unit of capacity*\n"
            "Kundur, Power System Stability and Control (1994)", size=18, lines=2)


def timeline(prs):
    s = add_slide(prs, "Who responds, and when", layout="Title and Content", footer=FOOTER)
    x0, w, t0, t1 = 3.55, 9.0, 0.1, 3600
    X = lambda t: x0 + w * (math.log10(t) - math.log10(t0)) / (math.log10(t1) - math.log10(t0))
    rows = (("Inertia", "kinetic energy, no control", t0, 10, "a2"),
            ("Primary", "governors sense frequency", 1, 30, "a5"),
            ("Secondary", "control restores 60 Hz", 30, 600, "a3"),
            ("Tertiary", "re-dispatch", 600, 3600, "a6"))
    y, h, gap = 1.3, 0.52, 0.13
    for i, (name, what, a, b, color) in enumerate(rows):
        yi = y + i * (h + gap)
        text(s, f"**{name}**", 0.75, yi, 2.7, h, size=22, anchor="m", name=f"{name} label")
        box(s, X(a), yi, X(b) - X(a), h, color, what, size=16, name=f"{name} span")
    axis_y = y + 4 * (h + gap) + 0.02
    arrow(s, (x0, axis_y), (x0 + w, axis_y), width=1.0, head=False, name="Time axis")
    for t, label in ((0.1, "0.1 s"), (1, "1 s"), (10, "10 s"), (60, "1 min"), (600, "10 min"), (3600, "1 hr")):
        arrow(s, (X(t), axis_y), (X(t), axis_y + 0.08), width=1.0, head=False, name="Tick")
        text(s, label, X(t) - 0.5, axis_y + 0.1, 1.0, 0.28, size=16, align="c", name="Tick label")
    bullets(s, [(0, "Each layer buys time for the next"),
                (0, "Only inertia acts with no sensing or control")],
            0.75, 4.75, 11.85, 1.3, anchor="t")
    note(s, "*Log time scale. Typical orders of magnitude; boundaries vary by interconnection*")


def model(prs):
    s = add_slide(prs, "The model in eight lines", footer=FOOTER)
    code(s, inspect.getsource(simulate).rstrip(), 0.95, 1.35, 11.7, 2.9)
    bullets(s, [(0, "Two states: frequency deviation and governor power"),
                (0, "Every number in this deck comes from this function")],
            0.75, 4.35, 11.85, 1.5, anchor="t")
    note(s, "*R = 20%: 5% droop on the quarter of capacity that responds. H and T are set per run*")


def chart_slide(prs):
    s = add_slide(prs, "Less inertia: a faster fall and a deeper low", footer=FOOTER)
    data = XyChartData()
    for name, run in (("H = 5 s", HIGH), ("H = 2 s", LOW)):
        ser = data.add_series(name)
        for t, f in run[run[:, 0] <= 40][::20]:
            ser.add_data_point(float(t), round(float(f), 4))
    frame = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, I(0.6), I(1.1), I(12.1), I(3.75), data)
    frame.name = "Frequency after the trip"
    ch = frame.chart
    style_chart(ch)
    fmin, fmax = 59.2, 60.0
    va, ca = ch.value_axis, ch.category_axis
    va.minimum_scale, va.maximum_scale, va.major_unit = fmin, fmax, 0.2
    ca.minimum_scale, ca.maximum_scale, ca.major_unit = 0, 40, 10
    va.tick_labels.number_format = "0.0"
    va.tick_labels.number_format_is_linked = False
    axis_title(va, "frequency (Hz)")
    axis_title(ca, "seconds after the trip")
    for ser, color in zip(ch.plots[0].series, ("a1", "a2")):
        ser.smooth = False
        paint(ser.format.line.color, color)
        ser.format.line.width = Pt(2.5)
    px, py, pw, ph = 1.6, 1.25, 10.9, 2.75
    fix_plot(frame, px, py, pw, ph)
    # Overlays are placed from the data, so they land exactly on the curves.
    X = lambda t: px + pw * t / 40
    Y = lambda f: py + ph * (fmax - f) / (fmax - fmin)
    ref = arrow(s, (X(0), Y(REFERENCE)), (X(40), Y(REFERENCE)), color="ink~", width=1.25, head=False, name="Reference line")
    ref.line.dash_style = MSO_LINE.DASH
    text(s, f"{REFERENCE} Hz reference", X(40) - 2.6, Y(REFERENCE) + 0.03, 2.6, 0.3, size=16, align="r", name="Reference label")
    for t, f, label, color, y in ((T_HIGH, F_HIGH, "H = 5 s", "a1", Y(REFERENCE) + 0.04), (T_LOW, F_LOW, "H = 2 s", "a2", Y(F_LOW) + 0.04)):
        circle(s, X(t), Y(f), 0.07, fill=color, line=None, name=f"Lowest point, {label}")
        text(s, f"[{color}-:**{label}**: low of {f:.2f} Hz at {t:.0f} s]", X(t) + 0.16, y, 4.2, 0.3, size=16,
             name=f"Low label, {label}")
    text(s, f"both settle at {SETTLE:.2f} Hz", X(40) - 3, Y(SETTLE) - 0.36, 3, 0.3, size=16, align="r", name="Settle label")
    bullets(s, [(0, "Same lost power, same final frequency"),
                (0, f"Steeper start ({slope(5):.2f} vs. {slope(2):.2f} Hz/s), and an earlier, deeper low")],
            0.75, 4.85, 11.85, 1.3, anchor="t")
    note(s, f"*Illustrative model: {LOSS:.0%} loss, 20% effective droop, governor lag {LAG} s*")


def tradeoff_picture(path, colors):
    Hs, Ts = np.linspace(1, 8, 22), np.linspace(1, 10, 22)
    low = np.array([[nadir(simulate(h, t, dt=0.02, t_end=40))[1] for t in Ts] for h in Hs])
    plt.rcParams.update({"font.family": ["Calibri", "Carlito", "DejaVu Sans"], "font.size": 16,
                         "text.color": colors["ink"], "axes.labelcolor": colors["ink"],
                         "axes.edgecolor": colors["ink"], "xtick.color": colors["ink"], "ytick.color": colors["ink"]})
    fig, ax = plt.subplots(figsize=(10.4, 3.5))
    ax.contourf(Ts, Hs, low, levels=[low.min() - 1, REFERENCE], colors=[colors["a2"]], alpha=0.16)
    thin = ax.contour(Ts, Hs, low, levels=[59.3, 59.4, 59.6, 59.7], colors=colors["ink"], linewidths=0.8, alpha=0.55)
    ax.clabel(thin, fmt="%.1f Hz", fontsize=14)
    bold = ax.contour(Ts, Hs, low, levels=[REFERENCE], colors=colors["a2"], linewidths=2.2)
    ax.clabel(bold, fmt="%.1f Hz", fontsize=15)
    for h, color in ((5, colors["a1"]), (2, colors["a2"])):
        ax.plot([LAG], [h], "o", color=color, ms=9, zorder=5)
        ax.annotate(f"H = {h} s", (LAG, h), (LAG + 0.15, h + 0.2), color=color, fontweight="bold", fontsize=16)
    ax.set_xlabel("governor lag T (s)")
    ax.set_ylabel("inertia H (s)")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    fig.tight_layout(pad=0.3)
    fig.savefig(path, dpi=300, transparent=True)
    plt.close(fig)


def tradeoff(prs):
    s = add_slide(prs, "Inertia and response speed trade off", footer=FOOTER)
    img = HERE / "_tradeoff.png"
    tradeoff_picture(img, theme_colors(prs))
    pic = s.shapes.add_picture(str(img), I(1.45), I(1.1), width=I(10.4))
    pic.name = "Lowest frequency vs. inertia and governor lag"
    pic._element.nvPicPr.cNvPr.set("descr", "Contours of the lowest frequency reached, by inertia and governor lag")
    img.unlink()
    bullets(s, [(0, "Lines show the lowest frequency reached"),
                (0, "Faster response can stand in for lost inertia")],
            0.75, 4.75, 11.85, 1.3, anchor="t")
    note(s, f"*{REFERENCE} Hz is a reference line; under-frequency load-shedding settings vary by region*")
    s.notes_slide.notes_text_frame.text = ("This plot is an image because PowerPoint has no contour chart. "
                                           "To change it, edit tradeoff_picture() in grid_frequency.py and re-run.")


def recap(prs):
    s = add_slide(prs, "Recap", layout="Title Only", footer=FOOTER)
    rows = [["Question", "Answer", "Example"],
            ["What sets frequency?", "Power balance", f"{LOSS:.0%} gap"],
            ["How fast does it fall?", "Gap ÷ inertia", f"{slope(5):.2f} vs. {slope(2):.2f} Hz/s"],
            ["How low does it go?", "Inertia vs. response speed", f"{F_HIGH:.2f} vs. {F_LOW:.2f} Hz"],
            ["Where does it settle?", "Droop and damping, not inertia", f"{SETTLE:.2f} Hz both"]]
    table(s, rows, 0.75, 2.0, [4.2, 4.6, 3.05], row_h=0.62, size=24, name="Recap")
    note(s, "*H = 5 s vs. H = 2 s, from the same illustrative model*")


def main(out):
    prs = Presentation(TEMPLATE)
    s = add_slide(prs, "Why grid frequency falls", layout="Title Slide")
    tf = s.placeholders[1].text_frame
    runs(tf.paragraphs[0], "Inertia, governors, and fewer spinning machines")
    p = tf.add_paragraph()
    p.space_before = Pt(40)
    runs(p, "An example deck", color="ink")
    questions(prs)
    balance(prs)
    swing(prs)
    questions(prs, dim=2)
    timeline(prs)
    model(prs)
    s = add_slide(prs, "Less inertia", layout="Section Header", footer=FOOTER)
    runs(s.placeholders[1].text_frame.paragraphs[0], "Same physics, less stored energy")
    chart_slide(prs)
    tradeoff(prs)
    recap(prs)
    prs.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(HERE / "grid_frequency.pptx"))
