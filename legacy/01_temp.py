from pathlib import Path
import subprocess
import tempfile

ROOT = Path(".")

docs = list(ROOT.rglob("*.doc"))

print(f"Found {len(docs)} .doc files")

for path in docs[:5]:

    print("\nTESTING:")
    print(path)

    with tempfile.TemporaryDirectory() as temp:

        result = subprocess.run(
            [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                temp,
                str(path),
            ],
            capture_output=True,
            text=True,
        )

        print("RETURN CODE:", result.returncode)
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)

        converted = Path(temp) / (path.stem + ".docx")

        print("CONVERTED:", converted.exists())