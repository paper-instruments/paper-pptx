#!/bin/zsh
# Render every deck in a folder with LibreOffice and compare the extracted text.
#
#   ./render_check.sh ~/Desktop/my-verification-set
#
# LibreOffice is a WEAK oracle: it is more permissive than PowerPoint and will happily
# render packages PowerPoint refuses. Use it for what it is good at -- proving that a deck
# which opens actually lays out to the SAME CONTENT as the control. A file that renders
# here has not been shown to be valid; a file that fails here is definitely broken.

set -e
DIR="${1:?usage: render_check.sh DIR}"
cd "$DIR"

command -v soffice >/dev/null 2>&1 || { echo "soffice not on PATH; skipping render"; exit 0; }

rm -rf _render && mkdir -p _render
for f in *.pptx; do
  soffice --headless --convert-to pdf --outdir _render "$f" >/dev/null 2>&1 || true
  base="${f%.pptx}"
  if [ -f "_render/$base.pdf" ]; then
    printf 'RENDERED  %-46s\n' "$f"
  else
    printf 'FAILED    %-46s  <- broken beyond even a permissive renderer\n' "$f"
  fi
done

echo
echo "--- extracted text, grouped by hash ---"
uv run --no-project --with pypdf python - <<'PY'
import hashlib
from pathlib import Path
from pypdf import PdfReader

groups = {}
for pdf in sorted(Path("_render").glob("*.pdf")):
    reader = PdfReader(str(pdf))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    # -- extracted text can carry lone surrogates from font/CMap quirks; hash the bytes
    # -- rather than failing, since only equality between decks matters here
    encoded = text.encode("utf-8", errors="replace")
    key = (len(reader.pages), hashlib.sha256(encoded).hexdigest()[:12])
    groups.setdefault(key, []).append(pdf.stem)

for (pages, digest), names in groups.items():
    print("  pages=%d text=%s" % (pages, digest))
    for name in names:
        print("      %s" % name)
print()
print("One group => every rendered deck shows identical content.")
print("More than one => a mutation changed what a renderer displays. Investigate that first.")
PY
