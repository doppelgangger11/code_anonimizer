import json
import pandas as pd
from pathlib import Path

from anonymizer.analyzer import analyze_project


BASE_DIR = Path('../' + input('>>> ')).resolve()
# OUTPUT_DIR = BASE_DIR / "COMPLETED"
# OUTPUT_DIR.mkdir(exist_ok=True)
MAPPING_FILE = BASE_DIR / "mapping.csv"

def main():
    project_dir = BASE_DIR

    files = analyze_project(project_dir)

    print()
    print("Project analysis")
    print("=" * 40)

    for file_type, paths in files.items():

        print(f"\n{file_type.upper()}: {len(paths)}")

        for path in paths:
            print(f"  {path.relative_to(project_dir)}")


if __name__ == "__main__":
    main()