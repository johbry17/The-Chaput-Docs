import csv
from pathlib import Path

ROOT = Path(".")
INPUT = ROOT / "raw_text" / "word_documents.csv"
OUTPUT = ROOT / "raw_text" / "top_100_documents.csv"

with INPUT.open(
    encoding="utf-8-sig",
    newline=""
) as f:
    records = list(csv.DictReader(f))

records.sort(
    key=lambda r: int(r["word_count"]),
    reverse=True
)

top = records[:100]

for i, record in enumerate(top, 1):
    record["rank"] = i

fieldnames = [
    "rank",
    "word_count",
    "path",
    "filename",
    "extension",
    "modified",
    "size_mb",
    "keyword_count",
    "keywords",
    "text_file",
    "preview",
]

with OUTPUT.open(
    "w",
    encoding="utf-8-sig",
    newline=""
) as f:

    writer = csv.DictWriter(
        f,
        fieldnames=fieldnames
    )

    writer.writeheader()
    writer.writerows(top)

print(f"Created: {OUTPUT}")