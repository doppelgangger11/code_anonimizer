import json
import pandas as pd
from tqdm import tqdm
from pathlib import Path

from anonymizer.analyzer import analyze_project
from anonymizer.excel import analyze_excel_files
from anonymizer.mapping_builder import (
    collect_columns,
    build_column_mapping,
    save_mapping,
    confirm_mappings
)


BASE_DIR = Path('../' + input('>>> ')).resolve()
# OUTPUT_DIR = BASE_DIR / "COMPLETED"
# OUTPUT_DIR.mkdir(exist_ok=True)
MAPPING_FILE = Path("./mapping.csv")

def main():
    project_dir = BASE_DIR

    # ========================================
    # PROJECT ANALYSIS
    # ========================================

    print()
    print("=" * 40)
    print("PROJECT ANALYSIS")
    print("=" * 40)

    files = analyze_project(project_dir)

    for file_type, paths in files.items():
        print(
            f"{file_type.upper():10} {len(paths)}"
        )

    # ========================================
    # EXCEL ANALYSIS
    # ========================================

    if files["excel"]:

        print()
        print("=" * 40)
        print("EXCEL ANALYSIS")
        print("=" * 40)

        excel_results = analyze_excel_files(
            files["excel"],
            project_dir,
        )

        print()
        print(
            f"✓ Excel analysis completed: "
            f"{len(excel_results)} files"
        )

    else:
        excel_results = {}

        print()
        print("No Excel files found.")

    print()

    for file_name, workbook in excel_results.items():

        for sheet_name, sheet in workbook["sheets"].items():
            print(file_name)

            print(
                "  └──"
                f" {sheet_name}: "
                f"header={sheet['header_row']}, "
                f"{sheet['rows']} rows, "
                f"{len(sheet['columns'])} columns"
            )
            # print(f"  columns:")
            # for column in sheet["columns"]:
            #     print(f"    - {column}")
            print('-' * 40)
            
    columns = collect_columns(excel_results)

    mappings = build_column_mapping(columns)

    mappings = confirm_mappings(mappings)

    save_mapping(
        mappings,
        MAPPING_FILE,
    )

    print()
    print(
        f"✓ Mapping saved: {MAPPING_FILE}"
    )


if __name__ == "__main__":
    main()