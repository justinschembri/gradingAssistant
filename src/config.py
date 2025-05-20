from pathlib import Path
ROOT_DIR = Path(__file__).parent.parent
WORKBOOK_DIR = ROOT_DIR / "src" / "workbooks"

if __name__ == "__main__":
    print(ROOT_DIR)
    print(WORKBOOK_DIR)
