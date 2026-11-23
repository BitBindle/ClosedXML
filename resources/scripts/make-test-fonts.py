# Make the TestFontStyles test fonts from the embedded CarlitoBare fonts.
# The family has only a regular and a bold style. The tests use it for text with a style that
# the family does not have. The fonts contain only the glyphs of the test text.
#
# Use: python3 make-test-fonts.py (needs fontTools, tested with 4.55.0)
from pathlib import Path

from fontTools import subset, ttLib

root = Path(__file__).resolve().parents[2]
source_dir = root / "ClosedXML" / "Graphics" / "Fonts"
target_dir = root / "ClosedXML.Tests" / "Resource" / "Fonts"
family = "TestFontStyles"

for style in ("Regular", "Bold"):
    font = ttLib.TTFont(source_dir / f"CarlitoBare-{style}.ttf")

    options = subset.Options()
    options.name_IDs = ["*"]
    options.notdef_outline = True
    subsetter = subset.Subsetter(options)
    subsetter.populate(text="Lorem ipsum dolor sit amet")
    subsetter.subset(font)

    # Record 0 is the copyright. Keep it.
    for rec in font["name"].names:
        if rec.nameID in (1, 16):
            rec.string = family
        elif rec.nameID == 4:
            rec.string = f"{family} {style}"
        elif rec.nameID in (3, 6):
            rec.string = f"{family}-{style}"

    font.save(target_dir / f"{family}-{style}.ttf")
