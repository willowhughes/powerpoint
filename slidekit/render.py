"""Render a .pptx to images so you can look at the result, and flag problems first.

    slidekit render deck.pptx [--out folder] [--engine powerpoint|libreoffice]

Writes slide-01.png, slide-02.png, ... and contact.png (every slide on one sheet). It uses
PowerPoint on Windows when installed, because that is what the audience sees, and LibreOffice
otherwise. Your own PowerPoint windows are left alone.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER as PH
from pptx.oxml.ns import qn

WIDTH, HEIGHT = 1600, 900

POWERPOINT = r"""
$ErrorActionPreference = 'Stop'
$running = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
try {{
  $p = $app.Presentations.Open('{src}', -1, 0, 0)
  foreach ($s in $p.Slides) {{ $s.Export(('{out}\slide-{{0:d2}}.png' -f $s.SlideIndex), 'PNG', {w}, {h}) }}
  $p.Close()
}} finally {{
  if (-not $running -and $app.Presentations.Count -eq 0) {{ $app.Quit() }}
}}
"""


def lint(path):
    """Catch what breaks a file or makes it hard to edit. Returns (errors, warnings)."""
    errors, warnings = [], []
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    for n, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if None in (sh.left, sh.top, sh.width, sh.height):
                continue
            if sh.width < 0 or sh.height < 0:
                errors.append(f"slide {n}: {sh.name!r} has a negative size, which PowerPoint reports as a corrupt file")
            elif sh.left < 0 or sh.top < 0 or sh.left + sh.width > W or sh.top + sh.height > H:
                warnings.append(f"slide {n}: {sh.name!r} runs off the slide")
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip():
                if sh.placeholder_format.type not in (PH.FOOTER, PH.SLIDE_NUMBER, PH.DATE):
                    warnings.append(f"slide {n}: empty placeholder {sh.name!r}; fill it or use another layout")
        el = slide._element
        hex_colors = sum(1 for _ in el.iter(qn("a:srgbClr")))
        fonts = sum(1 for f in el.iter(qn("a:latin")) if not f.get("typeface", "").startswith("+")
                    and f.get("typeface") not in ("Cambria Math", "Consolas"))
        if hex_colors:
            warnings.append(f"slide {n}: {hex_colors} hard-coded colors; they will not follow a template swap")
        if fonts:
            warnings.append(f"slide {n}: {fonts} hard-coded fonts; they will not follow a template swap")
    return errors, warnings


def has_powerpoint():
    if os.name != "nt":
        return False
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"))
        return True
    except OSError:
        return False


def find_soffice():
    known = ["/Applications/LibreOffice.app/Contents/MacOS/soffice",
             r"C:\Program Files\LibreOffice\program\soffice.exe"]
    return shutil.which("soffice") or shutil.which("libreoffice") or next((k for k in known if Path(k).exists()), None)


def via_powerpoint(src, out):
    quote = lambda p: str(p).replace("'", "''")
    script = POWERPOINT.format(src=quote(src), out=quote(out), w=WIDTH, h=HEIGHT)
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", script],
                       capture_output=True, text=True, timeout=300)
    if r.returncode or r.stderr.strip():
        raise RuntimeError("PowerPoint could not open or render the file, which usually means bad XML from "
                           f"the last edit:\n{r.stderr.strip().splitlines()[0]}")


def via_libreoffice(src, out):
    import pypdfium2 as pdfium
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("Rendering needs PowerPoint on Windows or LibreOffice (https://www.libreoffice.org).")
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, f"-env:UserInstallation={Path(tmp, 'profile').as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, capture_output=True, timeout=300)
        pdf = pdfium.PdfDocument(str(Path(tmp, src.stem + ".pdf")))
        for i in range(len(pdf)):
            page = pdf[i]
            page.render(scale=WIDTH / page.get_width()).to_pil().save(out / f"slide-{i + 1:02d}.png")
        pdf.close()


def contact_sheet(out, cols=3, thumb=640):
    pages = sorted(out.glob("slide-*.png"))
    if not pages:
        return
    h = round(thumb * HEIGHT / WIDTH)
    rows = -(-len(pages) // cols)
    sheet = Image.new("RGB", (cols * (thumb + 8) + 8, rows * (h + 8) + 8), (110, 110, 110))
    for i, page in enumerate(pages):
        x, y = 8 + (i % cols) * (thumb + 8), 8 + (i // cols) * (h + 8)
        sheet.paste(Image.open(page).convert("RGB").resize((thumb, h)), (x, y))
        ImageDraw.Draw(sheet).text((x + 4, y + 2), str(i + 1), fill=(200, 0, 0))
    sheet.save(out / "contact.png")


def render(src, out=None, engine=None):
    src = Path(src).resolve()
    out = Path(out).resolve() if out else src.with_name(src.stem + "_slides")
    errors, warnings = lint(src)
    for w in warnings:
        print("warning:", w)
    for e in errors:
        print("error:", e)
    if errors:
        return None
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("slide-*.png"):
        old.unlink()
    engine = engine or ("powerpoint" if has_powerpoint() else "libreoffice")
    (via_powerpoint if engine == "powerpoint" else via_libreoffice)(src, out)
    contact_sheet(out)
    print(f"rendered with {engine}: {out}")
    if engine == "libreoffice":
        print("note: LibreOffice is only an approximation of PowerPoint, so check fonts and equations in PowerPoint before presenting")
    return out


def compare(old, new):
    """Show what changed between two versions: changed slides side by side, old on the left."""
    with tempfile.TemporaryDirectory() as tmp:
        a, b = render(old, Path(tmp) / "a"), render(new, Path(tmp) / "b")
        if not a or not b:
            return None
        before, after = sorted(a.glob("slide-*.png")), sorted(b.glob("slide-*.png"))
        rows = []
        for i in range(max(len(before), len(after))):
            x = Image.open(before[i]).convert("RGB") if i < len(before) else None
            y = Image.open(after[i]).convert("RGB") if i < len(after) else None
            if x is None or y is None or ImageChops.difference(x, y).getbbox():
                rows.append((i + 1, x, y))
    new = Path(new).resolve()
    if not rows:
        print("No visible changes.")
        return new
    thumb, h = 800, 450
    sheet = Image.new("RGB", (2 * thumb + 24, len(rows) * (h + 8) + 8), (110, 110, 110))
    for r, (n, x, y) in enumerate(rows):
        for c, im in enumerate((x, y)):
            if im is not None:
                sheet.paste(im.resize((thumb, h)), (8 + c * (thumb + 8), 8 + r * (h + 8)))
        ImageDraw.Draw(sheet).text((12, 10 + r * (h + 8)), f"slide {n}", fill=(200, 0, 0))
    out = new.with_name(new.stem + "_changes.png")
    sheet.save(out)
    print(f"{len(rows)} of {max(len(before), len(after))} slides changed: {[n for n, _, _ in rows]}")
    print(f"Old on the left, new on the right: {out}")
    return out
