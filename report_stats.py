#!/usr/bin/env python3

from __future__ import annotations

import argparse
import io
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REQUIRED_COLUMNS = {
    "任务名称",
    "发表部门",
    "报表来源",
    "下发时间",
    "截止时间",
    "填报情况",
    "填报单位",
}
SOURCE_ORDER = ("区级", "市级", "省级")


@dataclass
class FillStats:
    report_count: int = 0
    completed_tasks: int = 0
    total_tasks: int = 0

    @property
    def rate(self) -> float:
        if self.total_tasks == 0:
            return 0.0
        return self.completed_tasks / self.total_tasks


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根据 Excel 表统计报表新增和填报率。")
    parser.add_argument("excel_path", help="Excel 文件路径（.xlsx）")
    parser.add_argument("--date", dest="anchor_date", help="统计所属周的任意日期，格式 YYYY-MM-DD")
    parser.add_argument("--start-date", help="自定义统计开始日期，格式 YYYY-MM-DD")
    parser.add_argument("--end-date", help="自定义统计结束日期，格式 YYYY-MM-DD")
    parser.add_argument("--sheet", help="工作表名称；不传时默认取第一个工作表")
    parser.add_argument("--save", help="把统计结果额外写入文本文件")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = generate_report_payload(
        excel_path=args.excel_path,
        anchor_date=args.anchor_date,
        start_date=args.start_date,
        end_date=args.end_date,
        sheet=args.sheet,
    )
    print(payload["output"])
    if args.save:
        Path(args.save).expanduser().resolve().write_text(payload["output"] + "\n", encoding="utf-8")
    return 0


def generate_report_payload(
    excel_path: str | Path,
    anchor_date: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    sheet: str | None = None,
    display_name: str | None = None,
) -> dict[str, object]:
    resolved_path = Path(excel_path).expanduser().resolve()
    if not resolved_path.exists():
        raise SystemExit(f"文件不存在: {resolved_path}")

    rows = load_rows(resolved_path, sheet)
    ensure_required_columns(rows)

    resolved_start, resolved_end = resolve_date_range(
        anchor_date=anchor_date,
        start_date=start_date,
        end_date=end_date,
    )
    summary = build_summary(rows, resolved_start, resolved_end)
    shown_name = display_name or str(resolved_path)
    output = render_output(summary, resolved_start, resolved_end, shown_name)
    return {
        "file_path": str(resolved_path),
        "display_name": shown_name,
        "start_date": resolved_start.isoformat(),
        "end_date": resolved_end.isoformat(),
        "summary": serialize_summary(summary),
        "output": output,
    }


def generate_circulation_export(
    excel_path: str | Path,
    anchor_date: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    sheet: str | None = None,
    display_name: str | None = None,
) -> tuple[str, list[dict[str, str]]]:
    payload = generate_report_payload(
        excel_path=excel_path,
        anchor_date=anchor_date,
        start_date=start_date,
        end_date=end_date,
        sheet=sheet,
        display_name=display_name,
    )
    rows = payload["summary"]["circulation_rows"]
    if not rows:
        raise SystemExit("当前统计周期内没有省级或市级报表流转明细。")
    return build_circulation_workbook(rows), rows


def load_rows(excel_path: Path, sheet_name: str | None) -> list[dict[str, str]]:
    with ZipFile(excel_path) as archive:
        shared_strings = load_shared_strings(archive)
        sheet_path = resolve_sheet_path(archive, sheet_name)
        sheet_root = ET.fromstring(archive.read(sheet_path))
        sheet_data = sheet_root.find("a:sheetData", NS)
        if sheet_data is None:
            return []

        matrix: list[list[str]] = []
        for row in sheet_data.findall("a:row", NS):
            values: dict[int, str] = {}
            for cell in row.findall("a:c", NS):
                ref = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)(\d+)", ref)
                if not match:
                    continue
                col_index = excel_col_to_number(match.group(1))
                values[col_index] = cell_value(cell, shared_strings)
            if values:
                max_col = max(values)
                matrix.append([values.get(i, "") for i in range(1, max_col + 1)])

    if not matrix:
        return []

    header = matrix[0]
    rows: list[dict[str, str]] = []
    for raw_row in matrix[1:]:
        padded = raw_row + [""] * (len(header) - len(raw_row))
        rows.append(dict(zip(header, padded)))
    return rows


def load_shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in item.iterfind(".//a:t", NS)))
    return values


def resolve_sheet_path(archive: ZipFile, sheet_name: str | None) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    sheets = workbook.find("a:sheets", NS)
    if sheets is None or not list(sheets):
        raise SystemExit("Excel 中没有可读取的工作表。")

    selected = None
    if sheet_name:
        for sheet in sheets:
            if sheet.attrib.get("name") == sheet_name:
                selected = sheet
                break
        if selected is None:
            raise SystemExit(f"未找到工作表: {sheet_name}")
    else:
        selected = list(sheets)[0]

    rid = selected.attrib.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    if not rid or rid not in rel_map:
        raise SystemExit("无法定位工作表文件。")
    return "xl/" + rel_map[rid].lstrip("/")


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iterfind(".//a:t", NS))

    raw_value = cell.find("a:v", NS)
    if raw_value is None or raw_value.text is None:
        return ""

    value = raw_value.text
    if cell_type == "s":
        return shared_strings[int(value)]
    return value


def excel_col_to_number(col: str) -> int:
    value = 0
    for ch in col:
        value = value * 26 + ord(ch.upper()) - 64
    return value


def ensure_required_columns(rows: list[dict[str, str]]) -> None:
    if not rows:
        raise SystemExit("Excel 中没有数据行。")

    missing = sorted(REQUIRED_COLUMNS - set(rows[0]))
    if missing:
        raise SystemExit(f"缺少必要列: {'、'.join(missing)}")


def resolve_date_range(
    anchor_date: str | None,
    start_date: str | None,
    end_date: str | None,
) -> tuple[date, date]:
    if start_date or end_date:
        if not start_date or not end_date:
            raise SystemExit("--start-date 和 --end-date 必须同时传入。")
        start = parse_input_date(start_date, "开始日期")
        end = parse_input_date(end_date, "结束日期")
        if start > end:
            raise SystemExit("开始日期不能晚于结束日期。")
        return start, end

    anchor = parse_input_date(anchor_date, "统计日期") if anchor_date else date.today()
    start = anchor - timedelta(days=anchor.weekday())
    end = start + timedelta(days=6)
    return start, end


def parse_input_date(value: str, field_name: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise SystemExit(f"{field_name}格式错误，应为 YYYY-MM-DD: {value}") from exc


def build_summary(rows: Iterable[dict[str, str]], start_date: date, end_date: date) -> dict[str, object]:
    new_rows = filter_rows_by_date(rows, "下发时间", start_date, end_date)
    due_rows = filter_rows_by_date(rows, "截止时间", start_date, end_date)

    new_counts = {source: 0 for source in SOURCE_ORDER}
    new_departments = {source: Counter() for source in SOURCE_ORDER}
    for row in new_rows:
        source = normalize_source(row.get("报表来源", ""))
        if source not in new_counts:
            continue
        new_counts[source] += 1
        department = row.get("发表部门", "").strip()
        if source == "区级":
            department = simplify_district_department(department)
        new_departments[source][department or "未填写部门"] += 1

    fill_counts = {source: FillStats() for source in SOURCE_ORDER}
    for row in due_rows:
        source = normalize_source(row.get("报表来源", ""))
        if source not in fill_counts:
            continue
        completed, total = parse_fill_stats(row.get("填报情况", ""))
        stats = fill_counts[source]
        stats.report_count += 1
        stats.completed_tasks += completed
        stats.total_tasks += total

    circulation_rows = []
    for row in new_rows:
        source = normalize_source(row.get("报表来源", ""))
        if source not in {"市级", "省级"}:
            continue
        circulation_rows.append(
            {
                "报表来源": source,
                "发表单位": row.get("发表部门", "").strip(),
                "发表时间": row.get("下发时间", "").strip(),
                "填表单位": row.get("填报单位", "").strip(),
            }
        )

    return {
        "new_counts": new_counts,
        "new_departments": new_departments,
        "fill_counts": fill_counts,
        "circulation_rows": circulation_rows,
    }


def filter_rows_by_date(
    rows: Iterable[dict[str, str]],
    column: str,
    start_date: date,
    end_date: date,
) -> list[dict[str, str]]:
    matched = []
    for row in rows:
        parsed = parse_sheet_date(row.get(column, ""))
        if parsed and start_date <= parsed <= end_date:
            matched.append(row)
    return matched


def parse_sheet_date(raw: str) -> date | None:
    text = str(raw or "").strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    if re.fullmatch(r"\d+(\.\d+)?", text):
        numeric = float(text)
        if numeric > 1000:
            return (datetime(1899, 12, 30) + timedelta(days=numeric)).date()
    return None


def normalize_source(raw: str) -> str:
    return str(raw).strip()


def simplify_district_department(raw: str) -> str:
    text = str(raw or "").strip()
    if not text:
        return ""
    return re.split(r"\s*[-－—]\s*", text, maxsplit=1)[0].strip()


def parse_fill_stats(raw: str) -> tuple[int, int]:
    match = re.search(r"(\d+)\s*/\s*(\d+)", str(raw))
    if not match:
        return 0, 0
    return int(match.group(1)), int(match.group(2))


def render_output(summary: dict[str, object], start_date: date, end_date: date, display_name: str) -> str:
    new_counts = summary["new_counts"]
    new_departments = summary["new_departments"]
    fill_counts = summary["fill_counts"]

    district_dept_text = format_department_counts(new_departments["区级"])
    city_dept_text = format_optional_department_counts(new_departments["市级"])
    province_dept_text = format_optional_department_counts(new_departments["省级"])

    district_fill = fill_counts["区级"]
    city_fill = fill_counts["市级"]
    province_fill = fill_counts["省级"]

    return "\n".join(
        [
            f"统计文件: {display_name}",
            f"统计周期: {start_date.isoformat()} 至 {end_date.isoformat()}",
            "",
            "【全区用户】本周无新增用户，现共有6662位用户；",
            "",
            "【报表新增】"
            f"统计本周系统内新增的流转报表数量，其中区级新增报表{new_counts['区级']}张"
            f"（{district_dept_text}）；接收市级制发报表{new_counts['市级']}张{city_dept_text}；"
            f"接收省级制发市区分发表{new_counts['省级']}张{province_dept_text}；",
            "",
            "【报表填报】"
            f"从用户实际填报情况统计本周截止的报表任务填报率，其中区级报表{district_fill.report_count}张，"
            f"接收市级报表{city_fill.report_count}张；接收省级制发市级分发报表{province_fill.report_count}张；"
            f"区级报表分解到个人后填报任务共{district_fill.total_tasks}个，"
            f"已完成填报任务{district_fill.completed_tasks}个，"
            f"区级报表填报率{format_rate(district_fill.rate)}；"
            f"市级报表分解到个人的填表任务共{city_fill.total_tasks}个，"
            f"已完成填报任务{city_fill.completed_tasks}个，"
            f"市级制发报表填报率{format_rate(city_fill.rate)}；"
            f"省级制发市级分发报表分解到个人的填表任务共{province_fill.total_tasks}个，"
            f"已完成填报任务{province_fill.completed_tasks}个，"
            f"省级制发市级分发报表填报率{format_rate(province_fill.rate)}；",
        ]
    )


def format_department_counts(counter: Counter[str]) -> str:
    if not counter:
        return "无"
    return "、".join(
        f"{department}{count}张"
        for department, count in counter.most_common()
    )


def format_optional_department_counts(counter: Counter[str]) -> str:
    if not counter:
        return ""
    return f"（{format_department_counts(counter)}）"


def format_rate(value: float) -> str:
    return f"{value * 100:.2f}%"


def serialize_summary(summary: dict[str, object]) -> dict[str, object]:
    new_departments = summary["new_departments"]
    fill_counts = summary["fill_counts"]
    return {
        "new_counts": dict(summary["new_counts"]),
        "new_departments": {
            source: [
                {"department": department, "count": count}
                for department, count in counter.most_common()
            ]
            for source, counter in new_departments.items()
        },
        "fill_counts": {
            source: {
                "report_count": stats.report_count,
                "completed_tasks": stats.completed_tasks,
                "total_tasks": stats.total_tasks,
                "rate": format_rate(stats.rate),
            }
            for source, stats in fill_counts.items()
        },
        "circulation_rows": list(summary["circulation_rows"]),
    }


def build_circulation_workbook(rows: list[dict[str, str]]) -> bytes:
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="流转明细" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        )
        archive.writestr("xl/styles.xml", build_styles_xml())
        archive.writestr("xl/worksheets/sheet1.xml", build_sheet_xml(rows))
    return buffer.getvalue()


def build_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="16"/>
      <color rgb="FFFFFFFF"/>
      <name val="Microsoft YaHei"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FF1F4E78"/>
      <name val="Microsoft YaHei"/>
    </font>
    <font>
      <sz val="11"/>
      <color rgb="FF17324D"/>
      <name val="Microsoft YaHei"/>
    </font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF2F75B5"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFEAF3FF"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFD8E9FB"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="3">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/>
    </border>
    <border>
      <left style="medium"/><right style="medium"/><top style="medium"/><bottom style="medium"/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
</styleSheet>"""


def build_sheet_xml(rows: list[dict[str, str]]) -> str:
    header = ["报表来源", "发表单位", "发表时间", "填表单位"]
    data_rows = [
        ["本周新增上级报表的流转情况", "", "", ""],
        header,
        *[
            [
                row.get("报表来源", ""),
                row.get("发表单位", ""),
                row.get("发表时间", ""),
                row.get("填表单位", ""),
            ]
            for row in rows
        ],
    ]

    row_xml_parts = []
    for idx, values in enumerate(data_rows, start=1):
        cell_parts = []
        for col_index, value in enumerate(values, start=1):
            if idx == 1 and col_index > 1:
                continue
            ref = f"{column_letter(col_index)}{idx}"
            if idx == 1:
                style_id = "1"
            elif idx == 2:
                style_id = "2"
            else:
                striped = (idx - 3) % 2 == 1
                if col_index in {1, 3}:
                    style_id = "5" if striped else "3"
                else:
                    style_id = "6" if striped else "4"
            cell_parts.append(
                f'<c r="{ref}" t="inlineStr" s="{style_id}"><is><t>{xml_escape(value)}</t></is></c>'
            )
        row_attrs = ' ht="30" customHeight="1"' if idx == 1 else ' ht="22" customHeight="1"' if idx == 2 else ' ht="20" customHeight="1"'
        row_xml_parts.append(f'<row r="{idx}"{row_attrs}>{"".join(cell_parts)}</row>')

    last_row = len(data_rows)

    return (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:D{last_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" state="frozen"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="24"/>'
        '<cols>'
        '<col min="1" max="1" width="12" customWidth="1"/>'
        '<col min="2" max="2" width="26" customWidth="1"/>'
        '<col min="3" max="3" width="21" customWidth="1"/>'
        '<col min="4" max="4" width="24" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(row_xml_parts)}</sheetData>'
        f'<autoFilter ref="A2:D{last_row}"/>'
        '<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>'
        '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '<pageSetup orientation="landscape" paperSize="9"/>'
        '</worksheet>'
    )


def column_letter(index: int) -> str:
    letters: list[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def xml_escape(value: str) -> str:
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


if __name__ == "__main__":
    raise SystemExit(main())
