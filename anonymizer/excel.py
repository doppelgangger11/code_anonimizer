from pathlib import Path

import pandas as pd
from tqdm import tqdm


SUPPORTED_EXCEL_EXTENSIONS = {
    ".xlsx",
    ".xls",
}

HEADER_SCAN_ROWS = 20

def detect_header_row(
    file_path: Path,
    sheet_name: str,
) -> int:
    """
    Detect the most likely header row in an Excel sheet.

    Only the first HEADER_SCAN_ROWS rows are inspected.

    Returns:
        Zero-based row index.
    """

    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=HEADER_SCAN_ROWS,
    )

    if preview.empty:
        return 0

    best_row = 0
    best_score = float("-inf")

    for row_index, row in preview.iterrows():

        values = [
            value
            for value in row.tolist()
            if pd.notna(value)
            and str(value).strip() != ""
        ]

        if not values:
            continue

        non_empty = len(values)

        unique_values = len(
            {
                str(value).strip().lower()
                for value in values
            }
        )

        uniqueness_ratio = unique_values / non_empty

        # How many cells are actually filled.
        density = non_empty / len(row)

        # Headers usually contain several unique values.
        if non_empty < 3:
            continue

        score = (
            non_empty * 3
            + density * 10
            + uniqueness_ratio * 5
        )

        if score > best_score:
            best_score = score
            best_row = row_index

    return best_row


def analyze_excel(file_path: Path) -> dict:
    """
    Analyze a single Excel workbook.

    Returns workbook structure:
    {
        "file": Path,
        "sheets": {
            "Sheet1": {
                "header_row": int,
                "columns": list,
                "rows": int
            }
        }
    }
    """

    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {file_path}"
        )

    if file_path.suffix.lower() not in SUPPORTED_EXCEL_EXTENSIONS:
        raise ValueError(
            f"Unsupported Excel format: {file_path.suffix}"
        )

    excel_file = pd.ExcelFile(file_path)

    sheets = {}

    for sheet_name in excel_file.sheet_names:
        
        header_row = detect_header_row(
            file_path,
            sheet_name,
        )

        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=header_row,
        )

        sheets[sheet_name] = {
            "header_row": header_row,
            "columns": df.columns.tolist(),
            "rows": len(df),
        }

    return {
        "file": file_path,
        "sheets": sheets,
    }


def analyze_excel_files(
    files: list[Path],
    project_dir: Path,
) -> dict:
    """
    Analyze multiple Excel files.
    """

    results = {}

    progress = tqdm(
        files,
        desc="Excel files",
        unit="file",
    )

    for file_path in progress:

        relative_path = file_path.relative_to(project_dir)

        progress.set_postfix_str(
            str(relative_path)
        )

        results[str(relative_path)] = analyze_excel(
            file_path
        )

    return results