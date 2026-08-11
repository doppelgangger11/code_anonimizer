import pandas as pd
from pathlib import Path


def collect_columns(excel_results: dict) -> list[str]:
    """
    Collect unique column names from analyzed Excel files.

    Returns:
        Sorted list of unique column names.
    """

    columns = set()

    for workbook in excel_results.values():
        for sheet in workbook["sheets"].values():
            for column in sheet["columns"]:

                if pd.isna(column):
                    continue

                column = str(column).strip()

                if not column:
                    continue

                columns.add(column)

    return sorted(columns)


def build_column_mapping(columns: list[str]) -> list[dict]:
    """
    Interactively build column mappings.
    """

    mappings = []

    print()
    print("=" * 40)
    print("COLUMN MAPPING")
    print("=" * 40)

    print()
    print(f"Found {len(columns)} unique columns.")
    print()

    for index, column in enumerate(columns, start=1):

        print(f"[{index}/{len(columns)}] {column}")

        replacement = input(
            "Replacement (empty = skip): "
        ).strip()

        if not replacement:
            print("  → skipped")
            print()
            continue

        mappings.append({
            "type": "column",
            "category": "column",
            "original": column,
            "replacement": replacement,
        })

        print(
            f"  → {column} -> {replacement}"
        )
        print()

    return mappings


def save_mapping(
    mappings: list[dict],
    output_path: Path,
) -> None:
    """
    Save mappings to CSV.
    """

    df = pd.DataFrame(
        mappings,
        columns=[
            "type",
            "category",
            "original",
            "replacement",
        ],
    )

    df.to_csv(
        output_path,
        index=False,
        encoding="utf-8-sig",
    )
    

def validate_mappings(
    mappings: list[dict],
) -> list[str]:
    """
    Validate generated mappings.

    Returns:
        List of validation errors.
    """

    errors = []

    # ----------------------------------------
    # Original -> replacement
    # ----------------------------------------

    seen_originals = set()

    for mapping in mappings:
        original = mapping["original"]

        if original in seen_originals:
            errors.append(
                f"Duplicate original value: {original}"
            )

        seen_originals.add(original)

    # ----------------------------------------
    # Column replacements
    # ----------------------------------------

    column_replacements = {}

    for mapping in mappings:

        if mapping["type"] != "column":
            continue

        replacement = mapping["replacement"]

        if replacement in column_replacements:
            previous = column_replacements[
                replacement
            ]

            errors.append(
                "Duplicate column replacement: "
                f"{previous} -> {replacement}, "
                f"{mapping['original']} -> {replacement}"
            )

        column_replacements[replacement] = (
            mapping["original"]
        )

    return errors


def print_mappings(
    mappings: list[dict],
) -> None:
    """
    Display current mappings.
    """

    print()
    print("=" * 40)
    print("CURRENT MAPPING")
    print("=" * 40)

    if not mappings:
        print("No mappings created.")
        return

    for index, mapping in enumerate(
        mappings,
        start=1,
    ):
        print(
            f"[{index}] "
            f"{mapping['original']} "
            f"-> "
            f"{mapping['replacement']}"
        )

    print("=" * 40)
    
    
def edit_mappings(
    mappings: list[dict],
) -> None:
    """
    Interactively edit existing mappings.
    """

    while True:

        print_mappings(mappings)

        choice = input(
            "Enter mapping number to edit "
            "(empty = finish): "
        ).strip()

        if not choice:
            break

        if not choice.isdigit():
            print("ERROR: Enter a number.")
            continue

        index = int(choice) - 1

        if index < 0 or index >= len(mappings):
            print("ERROR: Invalid mapping number.")
            continue

        mapping = mappings[index]

        print()
        print(
            f"Current: "
            f"{mapping['original']} "
            f"-> "
            f"{mapping['replacement']}"
        )

        replacement = input(
            "New replacement: "
        ).strip()

        if not replacement:
            print(
                "ERROR: Replacement cannot be empty."
            )
            continue

        mapping["replacement"] = replacement

        print(
            f"✓ Updated: "
            f"{mapping['original']} "
            f"-> "
            f"{replacement}"
        )

        errors = validate_mappings(mappings)

        if errors:
            print()
            print("VALIDATION ERRORS:")

            for error in errors:
                print(f"  ERROR: {error}")
        else:
            print("✓ Mapping is valid")
            

def confirm_mappings(
    mappings: list[dict],
) -> list[dict]:
    """
    Review and edit mappings until confirmed.
    """

    while True:

        print_mappings(mappings)

        errors = validate_mappings(mappings)

        if errors:
            print()
            print("VALIDATION ERRORS:")

            for error in errors:
                print(f"  ERROR: {error}")

            edit_mappings(mappings)
            continue

        answer = input(
            "\nEverything correct? [Y/n]: "
        ).strip().lower()

        if answer in ("", "y", "yes"):
            print()
            print("✓ Mapping confirmed")
            return mappings

        if answer in ("n", "no"):
            edit_mappings(mappings)
            continue

        print(
            "Please enter Y or N."
        )