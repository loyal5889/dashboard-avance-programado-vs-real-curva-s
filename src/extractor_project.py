"""Extrae las tareas de un único archivo MPP mediante Microsoft Project COM."""

from __future__ import annotations

import csv
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

try:
    import pythoncom
    import win32com.client
except ImportError as exc:  # mensaje más útil que el traceback de importación
    raise SystemExit(
        "ERROR: falta pywin32. Ejecute: python -m pip install -r requirements.txt"
    ) from exc


ROOT = Path(__file__).resolve().parents[1]
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "cronograma_darca2.csv"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

COLUMNS = [
    "ProjectName", "ExtractionDate", "StatusDate", "HoursPerDay", "HoursPerWeek",
    "MinutesPerDay", "MinutesPerWeek", "ID", "UniqueID", "WBS", "OutlineLevel",
    "TaskName", "Start", "Finish", "DurationMinutes", "DurationDays",
    "BaselineStart", "BaselineFinish", "BaselineDurationMinutes", "BaselineDurationDays",
    "PercentComplete", "PhysicalPercentComplete", "ActualStart", "ActualFinish",
    "ActualDurationMinutes", "ActualDurationDays", "RemainingDurationMinutes",
    "RemainingDurationDays", "Predecessors", "Successors", "TotalSlackMinutes",
    "TotalSlackDays", "FreeSlackMinutes", "FreeSlackDays", "Critical", "Summary",
    "Milestone", "Active", "ConstraintType", "ConstraintTypeName", "ConstraintDate",
    "Deadline", "WorkMinutes", "ActualWorkMinutes", "RemainingWorkMinutes",
    "BaselineWorkMinutes", "Cost", "BaselineCost", "BaselineAvailable",
]

CONSTRAINT_TYPES = {
    0: "As Soon As Possible", 1: "As Late As Possible", 2: "Must Start On",
    3: "Must Finish On", 4: "Start No Earlier Than", 5: "Start No Later Than",
    6: "Finish No Earlier Than", 7: "Finish No Later Than",
}

MISSING_TEXT_VALUES = {"NA", "N/A", "NOD", "NO DISPONIBLE", "NONE"}


def find_single_mpp(folder: Path) -> Path:
    files = sorted(folder.glob("*.mpp"))
    if len(files) != 1:
        raise RuntimeError(
            f"Se esperaba exactamente un archivo .mpp en {folder}; se encontraron {len(files)}."
        )
    return files[0]


def com_value(obj: Any, property_name: str, default: Any = "") -> Any:
    """Lee una propiedad COM; los campos no admitidos quedan vacíos."""
    try:
        value = getattr(obj, property_name)
    except Exception:
        return default
    if value is None:
        return default
    if isinstance(value, str) and value.strip().upper() in MISSING_TEXT_VALUES:
        return default
    return value


def format_date(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (datetime, date)):
        if isinstance(value, date) and not isinstance(value, datetime):
            value = datetime.combine(value, datetime.min.time())
        return value.strftime(DATE_FORMAT)
    text = str(value).strip()
    if not text or text.upper() in MISSING_TEXT_VALUES:
        return ""
    try:
        return datetime.fromisoformat(text).strftime(DATE_FORMAT)
    except ValueError:
        return text


def number(value: Any) -> Any:
    if value in (None, ""):
        return ""
    try:
        return float(value)
    except (TypeError, ValueError):
        return ""


def integer(value: Any) -> Any:
    parsed = number(value)
    return "" if parsed == "" else int(parsed)


def boolean(value: Any) -> Any:
    if value in (None, ""):
        return ""
    return bool(value)


def minutes_to_days(value: Any, minutes_per_day: float) -> Any:
    parsed = number(value)
    if parsed == "" or not minutes_per_day:
        return ""
    return parsed / minutes_per_day


def task_to_row(task: Any, project_name: str, extraction_date: str, status_date: str,
                hours_per_day: float, hours_per_week: float) -> dict[str, Any]:
    minutes_per_day = hours_per_day * 60
    minutes_per_week = hours_per_week * 60
    duration = com_value(task, "Duration")
    baseline_duration = com_value(task, "BaselineDuration")
    actual_duration = com_value(task, "ActualDuration")
    remaining_duration = com_value(task, "RemainingDuration")
    total_slack = com_value(task, "TotalSlack")
    free_slack = com_value(task, "FreeSlack")
    baseline_start = format_date(com_value(task, "BaselineStart"))
    baseline_finish = format_date(com_value(task, "BaselineFinish"))
    constraint_type = integer(com_value(task, "ConstraintType"))
    row = {
        "ProjectName": project_name,
        "ExtractionDate": extraction_date,
        "StatusDate": status_date,
        "HoursPerDay": hours_per_day,
        "HoursPerWeek": hours_per_week,
        "MinutesPerDay": minutes_per_day,
        "MinutesPerWeek": minutes_per_week,
        "ID": integer(com_value(task, "ID")),
        "UniqueID": integer(com_value(task, "UniqueID")),
        "WBS": str(com_value(task, "WBS")).strip(),
        "OutlineLevel": integer(com_value(task, "OutlineLevel")),
        "TaskName": str(com_value(task, "Name")).strip(),
        "Start": format_date(com_value(task, "Start")),
        "Finish": format_date(com_value(task, "Finish")),
        "DurationMinutes": number(duration),
        "DurationDays": minutes_to_days(duration, minutes_per_day),
        "BaselineStart": baseline_start,
        "BaselineFinish": baseline_finish,
        "BaselineDurationMinutes": number(baseline_duration),
        "BaselineDurationDays": minutes_to_days(baseline_duration, minutes_per_day),
        "PercentComplete": number(com_value(task, "PercentComplete")),
        "PhysicalPercentComplete": number(com_value(task, "PhysicalPercentComplete")),
        "ActualStart": format_date(com_value(task, "ActualStart")),
        "ActualFinish": format_date(com_value(task, "ActualFinish")),
        "ActualDurationMinutes": number(actual_duration),
        "ActualDurationDays": minutes_to_days(actual_duration, minutes_per_day),
        "RemainingDurationMinutes": number(remaining_duration),
        "RemainingDurationDays": minutes_to_days(remaining_duration, minutes_per_day),
        "Predecessors": str(com_value(task, "Predecessors")).strip(),
        "Successors": str(com_value(task, "Successors")).strip(),
        "TotalSlackMinutes": number(total_slack),
        "TotalSlackDays": minutes_to_days(total_slack, minutes_per_day),
        "FreeSlackMinutes": number(free_slack),
        "FreeSlackDays": minutes_to_days(free_slack, minutes_per_day),
        "Critical": boolean(com_value(task, "Critical")),
        "Summary": boolean(com_value(task, "Summary")),
        "Milestone": boolean(com_value(task, "Milestone")),
        "Active": boolean(com_value(task, "Active")),
        "ConstraintType": constraint_type,
        "ConstraintTypeName": CONSTRAINT_TYPES.get(constraint_type, ""),
        "ConstraintDate": format_date(com_value(task, "ConstraintDate")),
        "Deadline": format_date(com_value(task, "Deadline")),
        "WorkMinutes": number(com_value(task, "Work")),
        "ActualWorkMinutes": number(com_value(task, "ActualWork")),
        "RemainingWorkMinutes": number(com_value(task, "RemainingWork")),
        "BaselineWorkMinutes": number(com_value(task, "BaselineWork")),
        "Cost": number(com_value(task, "Cost")),
        "BaselineCost": number(com_value(task, "BaselineCost")),
        "BaselineAvailable": bool(baseline_start and baseline_finish),
    }
    return row


def write_csv(rows: Iterable[dict[str, Any]], destination: Path) -> int:
    rows = list(rows)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def extract(mpp_path: Path) -> tuple[int, int]:
    pythoncom.CoInitialize()
    app = None
    project_opened = False
    try:
        app = win32com.client.DispatchEx("MSProject.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.FileOpen(str(mpp_path))
        project_opened = True
        for _ in range(50):
            if app.Projects.Count > 0:
                break
            time.sleep(0.1)
        if app.Projects.Count == 0:
            raise RuntimeError(
                "Microsoft Project inició, pero no cargó el MPP. Ábralo manualmente para comprobar "
                "licencia, Vista protegida, contraseña o mensajes pendientes."
            )
        project = app.ActiveProject
        hours_per_day = float(com_value(project, "HoursPerDay", 0) or 0)
        hours_per_week = float(com_value(project, "HoursPerWeek", 0) or 0)
        if hours_per_day <= 0:
            raise RuntimeError("Project no devolvió un valor válido para HoursPerDay.")
        extraction_date = datetime.now().strftime(DATE_FORMAT)
        project_name = str(com_value(project, "Name") or mpp_path.stem).strip()
        status_date = format_date(com_value(project, "StatusDate"))
        rows = []
        for task in project.Tasks:
            if task is None:
                continue
            row = task_to_row(task, project_name, extraction_date, status_date,
                              hours_per_day, hours_per_week)
            if not row["TaskName"] and row["ID"] == "" and row["UniqueID"] == "":
                continue
            rows.append(row)
        count = write_csv(rows, OUTPUT_FILE)
        baseline_count = sum(bool(row["BaselineAvailable"]) for row in rows)
        return count, baseline_count
    finally:
        if app is not None:
            try:
                if project_opened:
                    app.FileClose(0)
            except Exception:
                pass
            try:
                app.Quit(0)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> int:
    print("Extractor DARCA 2 - Microsoft Project a CSV")
    try:
        mpp_path = find_single_mpp(INPUT_DIR)
        print(f"Archivo encontrado: {mpp_path}")
        count, baseline_count = extract(mpp_path)
        print(f"OK: {count} tareas exportadas a: {OUTPUT_FILE}")
        if baseline_count == 0:
            print("ADVERTENCIA: no se detectó una línea base válida en ninguna tarea.")
        else:
            print(f"Línea base detectada en {baseline_count} tareas.")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
