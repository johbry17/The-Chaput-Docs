# The Chaput Docs

> **Note**: The actual files, the Chaput Docs, are .gitignored. Dont be nosy; respect the deceased.

A personal project to rescue, organize, and make searchable a collection of old Chaput family files from three thumb drives.

The immediate goal: find a manuscript called "Walking on Air" and make the archive usable for my aunt without her having to manually open thousands of files.

---

## What's here

```
The Chaput Docs/
├── 1/                              # Original files — thumb drive 1
├── 2/                              # Original files — thumb drive 2
├── 3/                              # Original files — thumb drive 3
├── raw_text/                       # Extracted text + intermediate CSV catalogs
├── Chaput_Document_Archive.xlsx    # Word documents, sorted by word count
├── Chaput_Other_Documents.xlsx     # PDFs, TXT, RTF, ODT, HTML, etc.
├── Chaput_Image_Archive.xlsx       # Image files
├── Chaput_Audio_Archive.xlsx       # Audio files
├── How To - The Chaput Docs.pdf    # Instructions for my aunt
└── legacy/                         # Scripts and intermediate working material
```

**`1/`, `2/`, `3/`** are the original files exactly as copied from the thumb drives — duplicates, old versions, and all. About 4,236 files across 120 folders: 1,668 documents, 1,630 images, 196 audio files, 742 other.

**`raw_text/`** and **the four Excel files** are the finished product. The spreadsheets contain relative hyperlinks back to the originals. Keep `1/`, `2/`, `3/`, and `raw_text/` in the same folder as the spreadsheets or the links will break.

---

## What I did

### 1. Inventoried the archive — `00_chaput_docs.py`

First step: figure out what was there. The script recursively scanned the three source folders, classified files, counted extensions, identified potential writing candidates, and wrote a human-readable summary.

Output: `legacy/_archive_inventory/` — `file_inventory.csv`, `novel_candidates.csv`, `summary.txt`.

The summary was the most useful thing; it told me there were ~1,500 old `.doc` files, which meant I'd need a way to open them.

### 2. Tested LibreOffice conversion — `01_temp.py`

`python-docx` handles `.docx` but not the old binary `.doc` format. This was a small diagnostic script to confirm that LibreOffice could convert `.doc` → `.docx` headlessly. It could.

### 3. Searched for "Walking on Air" — `02_walking_on_air_search.py`

Searched both filenames and extracted document text for the phrase. Supported `.doc`, `.docx`, `.pdf`, `.txt`, `.rtf`, `.md`. Old `.doc` files were temporarily converted through LibreOffice first.

Output: `legacy/_walking_on_air_search/walking_on_air_results.csv`.

Didn't give a definitive answer immediately, which led to the broader cataloging effort.

### 4. Cataloged all Word documents — `03_catalog_docs.py`

The big step. Recursively found every `.doc` and `.docx`, extracted text (via the `.doc` → LibreOffice → `.docx` → `python-docx` pipeline), saved individual `.txt` files, and built a catalog with filename, path, modification date, size, word count, writing-keyword matches, and a text preview.

Output: `raw_text/text/`, `raw_text/word_documents.csv`, `raw_text/potential_writing.csv`.

### 5. Found the largest documents — `04_export_top_docs.py`

Sorted `word_documents.csv` by word count and exported the top 100.

Output: `raw_text/top_100_documents.csv`.

The results showed several large manuscripts — *SCREAM*, *STEADMAN*, *MAIN MOONLIGHT AND MELONS*, *SMOOTH AS SILK*, and others, most with multiple near-duplicate versions. In particular:

> **Rank #6 — `1\Toshiba Files\STEADMAN 1.doc` — ~122,700 words** — this is probably "Walking on Air."

### 6. Built the final Excel catalogs — `final.py`, `other_final.py`, `images_final.py`, `audio_final.py`

Converted the CSV catalogs into Excel workbooks with clickable relative links.

- `final.py` → `Chaput_Document_Archive.xlsx` (Word docs, sorted by word count, frozen header, filterable)
- `other_final.py` → `Chaput_Other_Documents.xlsx` (other document formats, by type then filename)
- `images_final.py` → `Chaput_Image_Archive.xlsx` (by image type then filename)
- `audio_final.py` → `Chaput_Audio_Archive.xlsx` (by original folder location — the grouping seemed more meaningful than metadata, which was unreliable)

---

## The `legacy/` folder

Once the archive was done, I moved the scripts out of the root so they wouldn't clutter the finished product.

```
legacy/
├── _archive_inventory/         # Output from 00_chaput_docs.py
├── _walking_on_air_search/     # Output from 01_walking_on_air_search.py
├── 00_chaput_docs.py
├── 01_temp.py
├── 02_walking_on_air_search.py
├── 03_catalog_docs.py
├── 04_export_top_docs.py
├── audio_final.py
├── images_final.py
├── other_final.py
└── final.py
```

I just numbered the main scripts roughly in chronological order. The numbering is mine; there was no formal naming convention at the time.

All the scripts were originally run from the **repository root**. If they ever need to be rerun they'd need to be run from the root again (or have their paths adjusted).
