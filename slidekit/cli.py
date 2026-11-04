"""The slidekit command: run deck scripts, render, inspect, back up, compare and restore."""
import argparse
import datetime
import os
import shutil
import subprocess
import sys
from pathlib import Path

def run(args):
    env = {**os.environ, "PYTHONUTF8": "1"}
    return subprocess.call([sys.executable, args.script, *args.rest], env=env)


def render_cmd(args):
    from .render import render
    try:
        return 0 if render(args.deck, args.out, args.engine) else 1
    except RuntimeError as e:
        print(e)
        return 1


def _versions(deck):
    deck = Path(deck)
    return deck.with_name(deck.stem + "_versions")


def backup(args):
    """Save a dated copy of a deck before changing it."""
    deck = Path(args.deck)
    folder = _versions(deck)
    folder.mkdir(exist_ok=True)
    dst = folder / f"{deck.stem}-{datetime.datetime.now():%Y%m%d-%H%M%S}{deck.suffix}"
    shutil.copy2(deck, dst)
    print(dst)
    return 0


def _newest(deck):
    deck = Path(deck)
    saved = [p for p in _versions(deck).glob(f"{deck.stem}-*{deck.suffix}") if "-undone-" not in p.name]
    return max(saved, key=lambda p: p.name) if saved else None


def restore(args):
    """Put back the newest backup, or a named one. What it replaces is kept as an "undone" copy."""
    deck = Path(args.deck)
    source = Path(args.version) if args.version else _newest(deck)
    if source is None or not source.exists():
        print(f"No backup found for {deck}")
        return 1
    undone = _versions(deck) / f"{deck.stem}-undone-{datetime.datetime.now():%Y%m%d-%H%M%S}{deck.suffix}"
    shutil.copy2(deck, undone)
    shutil.copy2(source, deck)
    print(f"Restored {deck} from {source}\nThe version it replaced is saved as {undone}")
    return 0


def compare_cmd(args):
    from .render import compare
    old, new = args.old, args.new
    if new is None:
        newest = _newest(args.old)
        if newest is None:
            print(f"No backup found for {args.old}")
            return 1
        old, new = str(newest), args.old
    try:
        return 0 if compare(old, new) else 1
    except RuntimeError as e:
        print(e)
        return 1


def _describe(shape, indent, lines):
    from pptx.enum.shapes import MSO_SHAPE_TYPE as T
    pos = ""
    if shape.left is not None:
        pos = f"  x={shape.left / 914400:.2f} y={shape.top / 914400:.2f} w={shape.width / 914400:.2f} h={shape.height / 914400:.2f}"
    kind = str(shape.shape_type).split(" ")[0].split(".")[-1].lower() if shape.shape_type is not None else "shape"
    if shape.is_placeholder:
        kind = "placeholder:" + str(shape.placeholder_format.type).split(" ")[0].lower()
    detail = ""
    if shape.is_placeholder and "SLIDE_NUMBER" in str(shape.placeholder_format.type):
        detail = ""
    elif shape.has_text_frame and shape.text_frame.text.strip():
        detail =" | ".join(p.text for p in shape.text_frame.paragraphs if p.text.strip())
    elif getattr(shape, "has_chart", False) and shape.has_chart:
        ch = shape.chart
        ser = ch.plots[0].series[0]
        detail = f"{ch.chart_type} series={[s.name for s in ch.plots[0].series]} first values={list(ser.values)[:4]}"
    elif getattr(shape, "has_table", False) and shape.has_table:
        detail = " / ".join(c.text_frame.text for c in shape.table.rows[0].cells)
    lines.append(f"{indent}{shape.name} ({kind}){pos}" + (f": {detail}" if detail else ""))
    if shape.shape_type == T.GROUP:
        for child in shape.shapes:
            _describe(child, indent + "  ", lines)


def inspect(args):
    """List a deck or template in text: layouts, theme colors, and every shape by name."""
    from . import A14, M, equations, open_deck, theme_colors
    prs = open_deck(args.deck)
    print(f"{args.deck}: {prs.slide_width / 914400:.2f} x {prs.slide_height / 914400:.2f} in, {len(prs.slides)} slides")
    print("theme:", " ".join(f"{k}={v}" for k, v in theme_colors(prs).items()))
    print("layouts:")
    for lay in prs.slide_layouts:
        print(f"  {lay.name}: " + ", ".join(
            f"{p.placeholder_format.idx}={str(p.placeholder_format.type).split(' ')[0].lower()}" for p in lay.placeholders))
    for n, slide in enumerate(prs.slides, 1):
        print(f"\nslide {n} [{slide.slide_layout.name}]")
        lines = []
        for shape in slide.shapes:
            _describe(shape, "  ", lines)
        print("\n".join(lines))
        for i, (wrap, x, y, w, h) in enumerate(equations(slide)):
            body = next(wrap.iter(f"{{{A14}}}m"))
            print(f"  equation {i}  x={x:.2f} y={y:.2f} w={w:.2f} h={h:.2f}: " + "".join(t.text or "" for t in body.iter(f"{{{M}}}t")))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            print("  notes:", slide.notes_slide.notes_text_frame.text)
    return 0


def main():
    ap = argparse.ArgumentParser(prog="slidekit", description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("run", help="run a deck script with slidekit and its dependencies available")
    p.add_argument("script")
    p.add_argument("rest", nargs=argparse.REMAINDER)
    p.set_defaults(func=run)
    p = sub.add_parser("render", help="check a deck and turn its slides into images")
    p.add_argument("deck")
    p.add_argument("--out")
    p.add_argument("--engine", choices=["powerpoint", "libreoffice"])
    p.set_defaults(func=render_cmd)
    p = sub.add_parser("inspect", help="list a deck or template as text")
    p.add_argument("deck")
    p.set_defaults(func=inspect)
    p = sub.add_parser("backup", help="save a dated copy of a deck before changing it")
    p.add_argument("deck")
    p.set_defaults(func=backup)
    p = sub.add_parser("restore", help="put back the newest backup of a deck")
    p.add_argument("deck")
    p.add_argument("version", nargs="?", help="a specific backup file")
    p.set_defaults(func=restore)
    p = sub.add_parser("compare", help="see changed slides side by side; with one file, against its newest backup")
    p.add_argument("old")
    p.add_argument("new", nargs="?")
    p.set_defaults(func=compare_cmd)
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.exit(args.func(args) or 0)


if __name__ == "__main__":
    main()
