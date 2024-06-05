from pathlib import Path
ROOT_DIR = Path(__file__).parent.parent
OUTPUT_PATH = ROOT_DIR / "src" / "outputs"

if __name__ == "__main__":
    print(ROOT_DIR)
    print(OUTPUT_PATH)
