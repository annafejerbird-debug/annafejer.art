from pathlib import Path
import sys

from pypdf import PdfReader, PdfWriter


def main() -> int:
    if len(sys.argv) != 4:
        print("Usage: split_portfolio_pages.py INPUT_PDF OUTPUT_DIR PREFIX", file=sys.stderr)
        return 2

    pdf_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    prefix = sys.argv[3]

    reader = PdfReader(str(pdf_path))
    for index, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        out_path = out_dir / f"{prefix}-page-{index:02d}.pdf"
        with out_path.open("wb") as handle:
            writer.write(handle)

    print(len(reader.pages))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
