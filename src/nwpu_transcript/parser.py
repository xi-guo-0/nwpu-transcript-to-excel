from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Tuple

import pdfplumber


def _clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _convert_chinese_semester(term: str) -> str | None:
    if not term:
        return None
    term = term.strip()
    if term in {"学期", "总学分绩点"}:
        return None
    match = re.match(r"^(\d{4})-(\d{4})([春夏秋冬])$", term)
    if not match:
        return None
    start, end, season = match.groups()
    mapping = {"秋": "1", "春": "2", "冬": "3", "夏": "3"}
    number = mapping.get(season)
    if not number:
        return None
    return f"{start}-{end}-{number}"


def _convert_english_semester(term: str) -> str | None:
    if not term:
        return None
    cleaned = term.replace("\n", " ").replace("\r", " ").strip()
    match = re.match(r"^(\d)(?:st|nd|rd|th)?\s*(\d{4}-\d{4})$", cleaned)
    if not match:
        return None
    order, years = match.groups()
    return f"{years}-{order}"


def parse_chinese(pdf_path: Path) -> List[Dict[str, str]]:
    """Parse the Chinese transcript PDF into records compatible with the template."""
    left_records: List[Dict[str, str]] = []
    right_records: List[Dict[str, str]] = []
    with pdfplumber.open(pdf_path) as pdf:
        table = pdf.pages[0].extract_table(
            table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 6,
            }
        )

    if not table:
        return []

    segments = (
        {"name": 0, "credit": 3, "score": 4, "category": 5, "semester": 8},
        {"name": 10, "credit": 15, "score": 16, "category": 17, "semester": 20},
    )

    skip_prefixes = ("姓名", "民族", "班级", "课程名称", "毕业设计", "应修总学分", "国家英语")

    for row in table:
        if not row:
            continue
        for indices, bucket in zip(segments, (left_records, right_records)):
            name_idx = indices["name"]
            if name_idx >= len(row):
                continue
            name_raw = row[name_idx]
            name = _clean_text(name_raw)
            if not name:
                continue
            if any(name.startswith(prefix) for prefix in skip_prefixes):
                continue

            semester_idx = indices["semester"]
            semester_raw = row[semester_idx] if semester_idx < len(row) else None
            semester = _convert_chinese_semester(_clean_text(semester_raw))
            if not semester:
                continue

            credit_idx = indices["credit"]
            score_idx = indices["score"]
            category_idx = indices["category"]

            credit = _clean_text(row[credit_idx]) if credit_idx < len(row) else ""
            score = _clean_text(row[score_idx]) if score_idx < len(row) else ""
            category = _clean_text(row[category_idx]) if category_idx < len(row) else ""

            record = {
                "课程名": name,
                "分数": score,
                "学分": credit,
                "学时": "",
                "学时单位": "学时",
                "课程类别": category,
                "学期": semester,
            }
            bucket.append(record)

    return left_records + right_records


def _parse_header_positions(header: List[str]) -> Tuple[Dict[str, int], Dict[str, int]]:
    def find_indices(start_idx: int) -> Dict[str, int]:
        positions: Dict[str, int] = {}
        cursor = start_idx
        for label in ("Course", "Credit", "Score", "Type", "Semester"):
            found = None
            for idx in range(cursor, len(header)):
                cell = header[idx]
                if cell and label in cell:
                    found = idx
                    break
            if found is None:
                raise ValueError("Header missing expected label")
            positions[label.lower()] = found
            cursor = found + 1
        return positions

    left_positions = find_indices(0)
    right_positions = find_indices(left_positions["semester"] + 1)
    return left_positions, right_positions


def _extract_english_record(row: List[str], positions: Dict[str, int]) -> Dict[str, str] | None:
    name_idx = positions["course"]
    if name_idx >= len(row):
        return None
    name = _clean_text(row[name_idx])
    if not name or name == "Course":
        return None

    semester_idx = positions["semester"]
    semester_raw = row[semester_idx] if semester_idx < len(row) else None
    semester = _convert_english_semester(_clean_text(semester_raw))
    if not semester:
        return None

    credit_idx = positions["credit"]
    score_idx = positions["score"]
    type_idx = positions["type"]

    credit = _clean_text(row[credit_idx]) if credit_idx < len(row) else ""
    score = _clean_text(row[score_idx]) if score_idx < len(row) else ""
    course_type = _clean_text(row[type_idx]) if type_idx < len(row) else ""

    return {
        "课程名": name,
        "分数": score,
        "学分": credit,
        "学时": "",
        "学时单位": "",
        "课程类别": course_type,
        "学期": semester,
    }


def parse_english(pdf_path: Path) -> List[Dict[str, str]]:
    """Parse the English transcript PDF into records compatible with the template."""
    records: List[Dict[str, str]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 6,
                }
            )
            for table in tables:
                if not table:
                    continue

                header_idx = None
                header = None
                for idx, row in enumerate(table):
                    header_text = "".join(cell or "" for cell in row)
                    if "Course" in header_text and "Semester" in header_text:
                        header_idx = idx
                        header = row
                        break

                if header_idx is None or header is None:
                    continue

                try:
                    left_positions, right_positions = _parse_header_positions(header)
                except ValueError:
                    continue
                left_bucket: List[Dict[str, str]] = []
                right_bucket: List[Dict[str, str]] = []
                for row in table[header_idx + 1 :]:
                    left_record = _extract_english_record(row, left_positions)
                    if left_record:
                        left_bucket.append(left_record)
                    right_record = _extract_english_record(row, right_positions)
                    if right_record:
                        right_bucket.append(right_record)
                records.extend(left_bucket)
                records.extend(right_bucket)
    return records

