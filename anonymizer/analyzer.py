import json
from pathlib import Path

from .excel import analyze_excel_files

with open('./config.json', 'r') as config_file:
    config = json.load(config_file)


SUPPORTED_EXTENSIONS = config["supported_extensions"]
EXCLUDED_DIRECTORIES = config["excluded_directories"]


def analyze_project(project_dir: Path) -> dict:
    """
    Scan project directory and classify supported files.

    Returns:
        {
            "python": [Path, ...],
            "notebook": [Path, ...],
            "excel": [Path, ...],
            "csv": [Path, ...],
        }
    """

    project_dir = Path(project_dir).resolve()

    if not project_dir.exists():
        raise FileNotFoundError(
            f"Project directory not found: {project_dir}"
        )

    if not project_dir.is_dir():
        raise NotADirectoryError(
            f"Path is not a directory: {project_dir}"
        )

    files = {
        "python": [],
        "notebook": [],
        "excel": [],
        "csv": [],
    }

    for path in project_dir.rglob("*"):

        if not path.is_file():
            continue

        if any(
            directory in EXCLUDED_DIRECTORIES
            for directory in path.parts
        ):
            continue

        if any(
            part.endswith("_anonymized")
            for part in path.parts
        ):
            continue

        extension = path.suffix.lower()

        file_type = SUPPORTED_EXTENSIONS.get(extension)

        if file_type is None:
            continue

        files[file_type].append(path)

    return files