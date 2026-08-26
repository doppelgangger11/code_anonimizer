#!/usr/bin/env python3
"""
project_anonymizer.py
======================

Анонимизатор проектов: excel-файлы + jupyter-ноутбуки.

Идея работы:
  1. Обходим корень проекта, находим все .xlsx/.xls файлы.
  2. Для текстовых колонок (по умолчанию — с высокой кардинальностью,
     т.е. похожих на имена/компании/id, а не на статусы вида "Да/Нет")
     собираем все уникальные значения и генерируем для них
     анонимные замены (Company_001, Client_002, ...).
     Числа/даты/формулы/форматирование ячеек не трогаем -> типы данных
     и вид таблиц сохраняются.
  3. Единый словарь замен (оригинал -> анонимный токен) применяется:
       - к самим excel-файлам (создаётся анонимизированная копия),
       - к .ipynb (код-ячейки, markdown-ячейки и ВЫВОДЫ ячеек —
         текстовые/HTML outputs, где значения часто "светятся" в
         распечатках датафреймов),
       - опционально к именам файлов/папок, если они содержат
         чувствительные значения (например "ACME_report.xlsx").
  4. Т.к. замена везде идёт по ОДНОМУ словарю подстрок, порядок
     фильтраций/сравнений в коде не меняется:
         df[df['company'] == 'ACME Corp']
     превращается в
         df[df['company'] == 'Company_001']
     и т.к. в анонимизированном excel в этой же колонке будет
     лежать 'Company_001' — код продолжает работать так же, как
     раньше (сравнивайте это со своим кодом — динамически
     сконструированные строки скрипт не поймает, см. ограничения
     внизу файла).

Ничего не изменяет "на месте": всегда пишет результат в отдельную
папку (--output), оригиналы не трогаются.

ВАЖНО: файл со словарём замен (anonymization_map.json) — это ключ,
который де-анонимизирует все данные обратно. Храните его отдельно от
анонимизированного проекта, не коммитьте в общий репозиторий/чат.

Использование
-------------

1) Сухой прогон — посмотреть, какие колонки скрипт considerит
   "чувствительными", ничего не меняя:

   python project_anonymizer.py analyze /path/to/project

2) Анонимизация:

   python project_anonymizer.py anonymize /path/to/project \
       --output /path/to/project_anonymized \
       --map-file /path/to/anonymization_map.json \
       --columns "Компания,Клиент,Менеджер" \
       --anonymize-filenames

   Если --columns не указан, применяется автоэвристика (кардинальность).
   Можно комбинировать: --columns форсирует колонки, --exclude-columns
   исключает их, даже если эвристика их бы выбрала.

3) Обратное восстановление ноутбука (для отладки, не обязательно):

   python project_anonymizer.py deanonymize /path/to/notebook_anon.ipynb \
       --map-file /path/to/anonymization_map.json \
       --output /path/to/notebook_restored.ipynb
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl

EXCEL_EXTS = {".xlsx", ".xlsm"}  # .xls (старый бинарный формат) openpyxl не пишет;
                                  # при необходимости конвертируйте в .xlsx заранее.
NOTEBOOK_EXT = ".ipynb"

# Папки, которые всегда пропускаются при обходе (и при анализе, и при
# анонимизации) — независимо от того, на каком уровне вложенности они
# встретились. Сравнение по имени папки, регистронезависимое.
DEFAULT_EXCLUDED_DIRS = {"backup", "archive"}


def is_excluded(path: Path, root: Path, excluded_dirs: set) -> bool:
    """True, если путь лежит внутри одной из исключённых папок (на любом
    уровне вложенности относительно root)."""
    if not excluded_dirs:
        return False
    try:
        rel_parts = path.relative_to(root).parts
    except ValueError:
        rel_parts = path.parts
    parent_parts = rel_parts[:-1]  # сама папка исключений, а не имя файла
    return any(part.lower() in excluded_dirs for part in parent_parts)

DEFAULT_STOPVALUES = {
    "", "-", "—", "n/a", "na", "нет данных", "да", "нет", "yes", "no",
    "true", "false", "unknown", "неизвестно",
}


# ---------------------------------------------------------------------------
# Утилиты
# ---------------------------------------------------------------------------

def sanitize_prefix(column_name: str) -> str:
    """Превращает имя колонки в безопасный префикс для токена замены."""
    name = re.sub(r"[^0-9A-Za-zА-Яа-яЁё]+", "_", str(column_name)).strip("_")
    if not name:
        name = "Value"
    return name[:30]


def looks_like_stopvalue(value: str) -> bool:
    v = value.strip().lower()
    if v in DEFAULT_STOPVALUES:
        return True
    if len(v) <= 1:
        return True
    # чистое число / дата в виде текста — не трогаем, оно не "имя"
    if re.fullmatch(r"[\d.,\-/: ]+", v):
        return True
    return False


def find_excel_files(root: Path, excluded_dirs: set = frozenset()) -> List[Path]:
    return sorted(
        p for p in root.rglob("*")
        if p.is_file() and p.suffix.lower() in EXCEL_EXTS
        and not is_excluded(p, root, excluded_dirs)
    )


def find_notebooks(root: Path, excluded_dirs: set = frozenset()) -> List[Path]:
    return sorted(
        p for p in root.rglob("*")
        if p.is_file() and p.suffix.lower() == NOTEBOOK_EXT
        and not is_excluded(p, root, excluded_dirs)
    )


# ---------------------------------------------------------------------------
# Шаг 1: анализ excel-файлов -> кандидаты на анонимизацию
# ---------------------------------------------------------------------------

class ColumnStats:
    __slots__ = ("values", "total", "file_sheet_col")

    def __init__(self):
        self.values: set = set()
        self.total: int = 0
        self.file_sheet_col: List[Tuple[str, str, str]] = []


def collect_column_stats(excel_files: List[Path]) -> Dict[str, ColumnStats]:
    """
    Собирает статистику по КОЛОНКАМ (ключ = имя колонки, т.е. заголовок
    из первой строки листа). Если одноимённые колонки встречаются в
    разных файлах/листах — статистика объединяется, это и нужно, чтобы
    эвристика кардинальности работала на всём проекте сразу.
    """
    stats: Dict[str, ColumnStats] = defaultdict(ColumnStats)

    for path in excel_files:
        try:
            wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        except Exception as e:
            print(f"  [!] Не смог открыть {path}: {e}", file=sys.stderr)
            continue

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                continue
            header = [str(h) if h is not None else f"col_{i}" for i, h in enumerate(header)]

            for row in rows:
                for col_name, cell_value in zip(header, row):
                    if not isinstance(cell_value, str):
                        continue
                    st = stats[col_name]
                    st.total += 1
                    if not looks_like_stopvalue(cell_value):
                        st.values.add(cell_value)
                        st.file_sheet_col.append((str(path), sheet_name, col_name))
        wb.close()

    return stats


def choose_sensitive_columns(
    stats: Dict[str, ColumnStats],
    force_include: Optional[set] = None,
    force_exclude: Optional[set] = None,
    cardinality_threshold: float = 0.5,
    min_unique: int = 2,
) -> List[str]:
    force_include = force_include or set()
    force_exclude = force_exclude or set()
    chosen = []
    for col, st in stats.items():
        if col in force_exclude:
            continue
        if col in force_include:
            chosen.append(col)
            continue
        if st.total == 0:
            continue
        ratio = len(st.values) / st.total
        if len(st.values) >= min_unique and ratio >= cardinality_threshold:
            chosen.append(col)
    return sorted(set(chosen))


# ---------------------------------------------------------------------------
# Шаг 2: построение словаря замен
# ---------------------------------------------------------------------------

def build_value_mapping(
    excel_files: List[Path],
    sensitive_columns: List[str],
) -> Tuple[Dict[str, str], List[dict]]:
    """
    Возвращает:
      value_map: {оригинальное_значение: анонимный_токен}   (глобально)
      audit: список записей для отчёта (файл/лист/колонка/было/стало)
    """
    sensitive_set = set(sensitive_columns)
    value_map: Dict[str, str] = {}
    counters: Dict[str, int] = defaultdict(int)
    audit: List[dict] = []

    for path in excel_files:
        try:
            wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        except Exception:
            continue
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                continue
            header = [str(h) if h is not None else f"col_{i}" for i, h in enumerate(header)]

            for row in rows:
                for col_name, cell_value in zip(header, row):
                    if col_name not in sensitive_set:
                        continue
                    if not isinstance(cell_value, str) or looks_like_stopvalue(cell_value):
                        continue
                    if cell_value in value_map:
                        continue
                    prefix = sanitize_prefix(col_name)
                    counters[prefix] += 1
                    token = f"{prefix}_{counters[prefix]:03d}"
                    value_map[cell_value] = token
                    audit.append({
                        "file": str(path),
                        "sheet": sheet_name,
                        "column": col_name,
                        "original": cell_value,
                        "anonymized": token,
                    })
        wb.close()

    return value_map, audit


# ---------------------------------------------------------------------------
# Шаг 3: применение словаря к excel-файлам
# ---------------------------------------------------------------------------

def anonymize_excel_file(src: Path, dst: Path, value_map: Dict[str, str]) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(src, data_only=False)  # полная загрузка -> сохранит формулы/формат
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value in value_map:
                    cell.value = value_map[cell.value]  # число/дата-форматирование ячейки не трогаем
    wb.save(dst)


# ---------------------------------------------------------------------------
# Шаг 4: применение словаря к .ipynb (код, markdown, outputs)
# ---------------------------------------------------------------------------

def compile_replacer(value_map: Dict[str, str]) -> Optional[re.Pattern]:
    if not value_map:
        return None
    # заменяем сначала более длинные значения, чтобы избежать
    # "порчи" при вложенных подстроках ("ACME" внутри "ACME Corp")
    keys = sorted(value_map.keys(), key=len, reverse=True)
    pattern = "|".join(re.escape(k) for k in keys)
    return re.compile(pattern)


def replace_text(text: str, value_map: Dict[str, str], pattern: re.Pattern) -> str:
    return pattern.sub(lambda m: value_map[m.group(0)], text)


def _replace_in_source(source, value_map, pattern):
    """source в .ipynb — это либо строка, либо список строк."""
    if isinstance(source, list):
        return [replace_text(line, value_map, pattern) for line in source]
    if isinstance(source, str):
        return replace_text(source, value_map, pattern)
    return source


def anonymize_notebook_file(src: Path, dst: Path, value_map: Dict[str, str]) -> None:
    pattern = compile_replacer(value_map)
    with open(src, "r", encoding="utf-8") as f:
        nb = json.load(f)

    if pattern is not None:
        for cell in nb.get("cells", []):
            if "source" in cell:
                cell["source"] = _replace_in_source(cell["source"], value_map, pattern)

            for output in cell.get("outputs", []) or []:
                if "text" in output:
                    output["text"] = _replace_in_source(output["text"], value_map, pattern)
                data = output.get("data")
                if isinstance(data, dict):
                    for mime, payload in list(data.items()):
                        if mime.startswith("text/"):  # text/plain, text/html — не бинарные
                            data[mime] = _replace_in_source(payload, value_map, pattern)

    dst.parent.mkdir(parents=True, exist_ok=True)
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(nb, f, ensure_ascii=False, indent=1)


# ---------------------------------------------------------------------------
# Шаг 5 (опционально): анонимизация имён файлов/папок
# ---------------------------------------------------------------------------

def anonymize_path_name(name: str, value_map: Dict[str, str], pattern: Optional[re.Pattern]) -> str:
    if pattern is None:
        return name
    return replace_text(name, value_map, pattern)


# ---------------------------------------------------------------------------
# Оркестрация
# ---------------------------------------------------------------------------

def cmd_analyze(args):
    root = Path(args.root)
    excluded_dirs = set(x.strip().lower() for x in args.exclude_dirs.split(",") if x.strip())
    print(f"Исключённые папки: {sorted(excluded_dirs) or 'нет'}")
    excel_files = find_excel_files(root, excluded_dirs)
    print(f"Найдено excel-файлов: {len(excel_files)}")
    stats = collect_column_stats(excel_files)
    chosen = choose_sensitive_columns(
        stats,
        cardinality_threshold=args.cardinality_threshold,
    )
    print("\nКолонки-кандидаты на анонимизацию (эвристика по кардинальности):")
    for col in chosen:
        st = stats[col]
        ratio = len(st.values) / st.total if st.total else 0
        print(f"  - {col!r}: {len(st.values)} уникальных / {st.total} значений (ratio={ratio:.2f})")
    print("\nОстальные текстовые колонки (НЕ будут анонимизированы по умолчанию):")
    for col, st in stats.items():
        if col in chosen:
            continue
        ratio = len(st.values) / st.total if st.total else 0
        print(f"  - {col!r}: {len(st.values)} уникальных / {st.total} значений (ratio={ratio:.2f})")
    print(
        "\nЕсли эвристика выбрала не то — используйте --columns / --exclude-columns "
        "в команде anonymize."
    )


def cmd_anonymize(args):
    root = Path(args.root)
    output_root = Path(args.output) if args.output else root.parent / (root.name + "_anonymized")
    map_file = Path(args.map_file) if args.map_file else output_root.parent / (root.name + "_anonymization_map.json")

    force_include = set(x.strip() for x in args.columns.split(",")) if args.columns else set()
    force_exclude = set(x.strip() for x in args.exclude_columns.split(",")) if args.exclude_columns else set()
    excluded_dirs = set(x.strip().lower() for x in args.exclude_dirs.split(",") if x.strip())

    excel_files = find_excel_files(root, excluded_dirs)
    notebooks = find_notebooks(root, excluded_dirs)
    print(f"Исключённые папки: {sorted(excluded_dirs) or 'нет'}")
    print(f"Excel-файлов: {len(excel_files)}, ноутбуков: {len(notebooks)}")

    stats = collect_column_stats(excel_files)
    sensitive_columns = choose_sensitive_columns(
        stats,
        force_include=force_include,
        force_exclude=force_exclude,
        cardinality_threshold=args.cardinality_threshold,
    )
    print(f"Анонимизируемые колонки: {sensitive_columns}")

    value_map, audit = build_value_mapping(excel_files, sensitive_columns)
    print(f"Уникальных значений для замены: {len(value_map)}")

    # -- копируем всё дерево проекта как есть --
    if output_root.exists():
        if args.overwrite:
            shutil.rmtree(output_root)
        else:
            print(f"[!] {output_root} уже существует. Используйте --overwrite.", file=sys.stderr)
            sys.exit(1)
    def ignore_excluded(dir_path, names):
        # shutil.copytree's ignore callback: возвращаем имена, которые
        # НЕ нужно копировать. Так backup/archive не попадут в
        # анонимизированный вывод вообще — иначе там остались бы
        # неанонимизированные исходники, что сводит на нет весь смысл.
        if not excluded_dirs:
            return set()
        return {n for n in names if n.lower() in excluded_dirs}

    shutil.copytree(root, output_root, ignore=ignore_excluded)
    if excluded_dirs:
        print(f"  (папки {sorted(excluded_dirs)} пропущены целиком, в вывод не копировались)")

    # -- перезаписываем excel анонимизированными версиями --
    for src in excel_files:
        rel = src.relative_to(root)
        dst = output_root / rel
        anonymize_excel_file(src, dst, value_map)

    # -- перезаписываем ноутбуки --
    for src in notebooks:
        rel = src.relative_to(root)
        dst = output_root / rel
        anonymize_notebook_file(src, dst, value_map)

    filename_map = {}
    if args.anonymize_filenames:
        pattern = compile_replacer(value_map)
        if pattern is not None:
            # переименовываем файлы/папки снизу вверх, чтобы не сломать пути
            all_paths = sorted(output_root.rglob("*"), key=lambda p: len(p.parts), reverse=True)
            for p in all_paths:
                new_name = anonymize_path_name(p.name, value_map, pattern)
                if new_name != p.name:
                    new_path = p.with_name(new_name)
                    p.rename(new_path)
                    filename_map[str(p.relative_to(output_root))] = str(new_path.relative_to(output_root))

    # -- сохраняем словарь-ключ отдельно от анонимизированного проекта --
    map_file.parent.mkdir(parents=True, exist_ok=True)
    with open(map_file, "w", encoding="utf-8") as f:
        json.dump({
            "values": value_map,
            "filenames": filename_map,
            "audit": audit,
        }, f, ensure_ascii=False, indent=2)

    print(f"\nГотово.")
    print(f"  Анонимизированный проект: {output_root}")
    print(f"  Словарь-ключ (ХРАНИТЬ ОТДЕЛЬНО, НЕ ПУБЛИКОВАТЬ): {map_file}")


def cmd_deanonymize(args):
    with open(args.map_file, "r", encoding="utf-8") as f:
        m = json.load(f)
    reverse_map = {v: k for k, v in m["values"].items()}
    anonymize_notebook_file(Path(args.notebook), Path(args.output), reverse_map)
    print(f"Восстановлено: {args.output}")


def build_arg_parser():
    p = argparse.ArgumentParser(description="Анонимизатор excel + jupyter проектов")
    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("analyze", help="Сухой прогон: показать колонки-кандидаты")
    a.add_argument("root")
    a.add_argument("--cardinality-threshold", type=float, default=0.5)
    a.add_argument(
        "--exclude-dirs", default="backup,archive",
        help="Через запятую: папки, полностью пропускаемые при обходе (по умолчанию: backup,archive)",
    )
    a.set_defaults(func=cmd_analyze)

    b = sub.add_parser("anonymize", help="Выполнить анонимизацию")
    b.add_argument("root")
    b.add_argument("--output", help="Куда писать анонимизированный проект")
    b.add_argument("--map-file", help="Куда писать словарь-ключ (JSON)")
    b.add_argument("--columns", help="Через запятую: форсировать эти колонки")
    b.add_argument("--exclude-columns", help="Через запятую: исключить эти колонки")
    b.add_argument("--cardinality-threshold", type=float, default=0.5)
    b.add_argument("--anonymize-filenames", action="store_true")
    b.add_argument("--overwrite", action="store_true")
    b.add_argument(
        "--exclude-dirs", default="backup,archive",
        help="Через запятую: папки, полностью пропускаемые при обходе и НЕ копируемые в вывод "
             "(по умолчанию: backup,archive)",
    )
    b.set_defaults(func=cmd_anonymize)

    c = sub.add_parser("deanonymize", help="Восстановить один ноутбук по словарю (для отладки)")
    c.add_argument("notebook")
    c.add_argument("--map-file", required=True)
    c.add_argument("--output", required=True)
    c.set_defaults(func=cmd_deanonymize)

    return p


def main():
    parser = build_arg_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
