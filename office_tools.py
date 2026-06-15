from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


@dataclass
class TaskResult:
    message: str
    output_path: Path | None = None


def is_supported_file(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_EXTENSIONS


def run_office_task(file_path: Path, instruction: str, output_dir: Path) -> TaskResult:
    instruction_normalized = _normalize(instruction)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not is_supported_file(file_path):
        return TaskResult("Bu fayl tipi hələ dəstəklənmir. Hazırda `.xlsx`, `.xls`, `.csv` qəbul edirəm.")

    if _wants_columns(instruction_normalized):
        df = _read_table(file_path)
        columns = "\n".join(f"- {column}" for column in df.columns)
        return TaskResult(f"Fayldakı sütunlar:\n{columns}")

    if _wants_summary(instruction_normalized):
        return TaskResult(_summarize(file_path))

    if _wants_excel(instruction_normalized):
        df = _read_table(file_path)
        output_path = output_dir / f"{file_path.stem}_converted.xlsx"
        df.to_excel(output_path, index=False)
        return TaskResult("Faylı Excel formatına çevirdim.", output_path)

    if _wants_csv(instruction_normalized):
        df = _read_table(file_path)
        output_path = output_dir / f"{file_path.stem}_converted.csv"
        df.to_csv(output_path, index=False)
        return TaskResult("Faylı CSV formatına çevirdim.", output_path)

    filter_result = _try_filter(file_path, instruction, instruction_normalized, output_dir)
    if filter_result:
        return filter_result

    sort_result = _try_sort(file_path, instruction, instruction_normalized, output_dir)
    if sort_result:
        return sort_result

    return TaskResult(
        "Tapşırığı tam başa düşmədim. Belə yaza bilərsiniz: "
        "`xülasə ver`, `sütunları göstər`, `status = ödənilib filtr et`, "
        "`məbləğ sütununa görə sırala`, `csv et`, `excel et`."
    )


def _read_table(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_path)
    return pd.read_excel(file_path)


def _summarize(file_path: Path) -> str:
    df = _read_table(file_path)
    lines = [
        f"Fayl: `{file_path.name}`",
        f"Sətir sayı: {len(df)}",
        f"Sütun sayı: {len(df.columns)}",
        "",
        "Sütunlar:",
    ]
    lines.extend(f"- {column}" for column in df.columns)

    numeric_columns = df.select_dtypes(include="number").columns
    if len(numeric_columns) > 0:
        lines.append("")
        lines.append("Rəqəmsal sütunların qısa statistikası:")
        stats = df[numeric_columns].describe().round(2).transpose()
        for column, row in stats.iterrows():
            lines.append(
                f"- {column}: min={row['min']}, max={row['max']}, orta={row['mean']}"
            )

    missing = df.isna().sum()
    missing = missing[missing > 0]
    if not missing.empty:
        lines.append("")
        lines.append("Boş xanalar:")
        lines.extend(f"- {column}: {count}" for column, count in missing.items())

    return "\n".join(lines)


def _try_sort(
    file_path: Path,
    instruction: str,
    instruction_normalized: str,
    output_dir: Path,
) -> TaskResult | None:
    if "sirala" not in instruction_normalized and "sort" not in instruction_normalized:
        return None

    df = _read_table(file_path)
    column = _find_column(df, instruction)
    if column is None:
        return TaskResult("Sıralamaq üçün sütunu tapa bilmədim. Məsələn: `məbləğ sütununa görə sırala`.")

    descending = any(word in instruction_normalized for word in ["azalan", "boyukden", "desc"])
    sorted_df = df.sort_values(by=column, ascending=not descending)
    output_path = output_dir / f"{file_path.stem}_sorted.xlsx"
    sorted_df.to_excel(output_path, index=False)
    return TaskResult(f"`{column}` sütununa görə sıraladım.", output_path)


def _try_filter(
    file_path: Path,
    instruction: str,
    instruction_normalized: str,
    output_dir: Path,
) -> TaskResult | None:
    if "filtr" not in instruction_normalized and "filter" not in instruction_normalized:
        return None

    df = _read_table(file_path)
    match = re.search(r"(.+?)\s*(=|==|:)\s*(.+?)(?:\s+olanlari|\s+filtr|$)", instruction, flags=re.I)
    if not match:
        return TaskResult("Filtr şərtini tapa bilmədim. Məsələn: `status = ödənilib filtr et`.")

    raw_column, _, raw_value = match.groups()
    column = _find_column(df, raw_column)
    if column is None:
        return TaskResult("Filtr üçün sütunu tapa bilmədim. Sütun adını fayldakı kimi yazmağa çalışın.")

    value = raw_value.strip().strip("'\"")
    filtered = df[df[column].astype(str).str.casefold() == value.casefold()]
    output_path = output_dir / f"{file_path.stem}_filtered.xlsx"
    filtered.to_excel(output_path, index=False)
    return TaskResult(f"`{column}` = `{value}` olan {len(filtered)} sətir tapdım.", output_path)


def _find_column(df: pd.DataFrame, text: str) -> str | None:
    text_normalized = _normalize(text)
    normalized_columns = {_normalize(str(column)): str(column) for column in df.columns}

    for normalized, original in normalized_columns.items():
        if normalized and normalized in text_normalized:
            return original

    for normalized, original in normalized_columns.items():
        if normalized and text_normalized in normalized:
            return original

    return None


def _wants_summary(instruction: str) -> bool:
    return any(word in instruction for word in ["xulase", "summary", "analiz", "hesabat"])


def _wants_columns(instruction: str) -> bool:
    return any(phrase in instruction for phrase in ["sutun", "column", "basliq"])


def _wants_csv(instruction: str) -> bool:
    return "csv" in instruction


def _wants_excel(instruction: str) -> bool:
    return any(word in instruction for word in ["excel", "xlsx"])


def _normalize(text: str) -> str:
    replacements = {
        "ə": "e",
        "ö": "o",
        "ü": "u",
        "ı": "i",
        "ğ": "g",
        "ş": "s",
        "ç": "c",
        "Ə": "e",
        "Ö": "o",
        "Ü": "u",
        "I": "i",
        "İ": "i",
        "Ğ": "g",
        "Ş": "s",
        "Ç": "c",
    }
    lowered = text.casefold()
    for original, replacement in replacements.items():
        lowered = lowered.replace(original, replacement)
    return lowered

