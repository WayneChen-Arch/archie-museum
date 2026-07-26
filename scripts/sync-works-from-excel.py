#!/usr/bin/env python3
"""Generate or apply works.xlsx <-> index.html gallery copy.

Usage:
  python3 scripts/sync-works-from-excel.py export   # index.html -> works.xlsx
  python3 scripts/sync-works-from-excel.py import   # works.xlsx -> index.html
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from opencc import OpenCC

ROOT = Path(__file__).resolve().parents[1]
HTML_PATH = ROOT / "index.html"
XLSX_PATH = ROOT / "works.xlsx"

HEADERS = [
    "序号",
    "图片文件",
    "名称",
    "年份",
    "创作媒介",
    "一句话介绍",
    "创作过程说明",
    "名称（英文）",
    "创作媒介（英文）",
    "一句话介绍（英文）",
    "创作过程说明（英文）",
]

EDITABLE_NOTE = (
    "使用说明：\n"
    "1. 请主要修改「名称」「年份」「创作媒介」「一句话介绍」「创作过程说明」这些简体中文列。\n"
    "2. 「序号」和「图片文件」请勿改动，用于对应网页中的作品。\n"
    "3. 英文列可按需修改；若英文列留空，同步时会保留网页里原来的英文。\n"
    "4. 繁体中文无需手填：上传本文件到 GitHub 后，同步脚本会按简体自动转换为繁体，并写入网页三种语言。\n"
    "5. 修改完成后，把本文件上传/覆盖到仓库根目录的 works.xlsx，并告知需要同步到网页。"
)


def s2t(text: str) -> str:
    if not text:
        return ""
    return OpenCC("s2t").convert(text)


def extract_works_block(html: str) -> tuple[str, str, str]:
    start = html.find("const works = [")
    if start < 0:
        raise RuntimeError("找不到 const works = [")
    end = html.find("\n      ];", start)
    if end < 0:
        raise RuntimeError("找不到 works 数组结尾")
    end += len("\n      ];")
    return html[:start], html[start:end], html[end:]


def parse_works(html: str) -> list[dict]:
    """Parse works objects with a light, structured regex (stable for this file)."""
    _, block, _ = extract_works_block(html)
    # Split on top-level object starts after "const works = ["
    body = block[len("const works = [") : -len("];")].strip()
    if body.endswith(","):
        body = body[:-1]

    objs: list[str] = []
    depth = 0
    start = None
    for i, ch in enumerate(body):
        if ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start is not None:
                objs.append(body[start : i + 1])
                start = None

    works = []
    for obj in objs:
        work = {
            "titles": _pick_i18n(obj, "titles"),
            "year": _pick_str(obj, "year"),
            "materials": _pick_i18n(obj, "materials"),
            "image_file": _pick_art(obj, "image"),
            "size": _pick_str(obj, "size"),
            "notes": _pick_i18n(obj, "notes"),
            "process_file": _pick_art(obj, "processImage"),
            "processCaptions": _pick_i18n(obj, "processCaptions") or {},
            "series": _pick_str(obj, "series"),
            "panel": _pick_str(obj, "panel"),
            "raw": obj,
        }
        works.append(work)
    return works


def _pick_str(obj: str, key: str) -> str:
    m = re.search(rf'{key}:\s*"((?:\\.|[^"\\])*)"', obj)
    return _unescape(m.group(1)) if m else ""


def _pick_art(obj: str, key: str) -> str:
    m = re.search(rf'{key}:\s*art\("((?:\\.|[^"\\])*)"\)', obj)
    return _unescape(m.group(1)) if m else ""


def _pick_i18n(obj: str, key: str) -> dict:
    m = re.search(rf"{key}:\s*\{{([^{{}}]+)\}}", obj, re.S)
    if not m:
        return {}
    body = m.group(1)
    out = {}
    for lang in ("zh-Hans", "zh-Hant", "en"):
        lm = re.search(
            rf'(?:"{lang}"|{re.escape(lang)}):\s*"((?:\\.|[^"\\])*)"',
            body,
        )
        if lm:
            out[lang] = _unescape(lm.group(1))
    return out


def _unescape(s: str) -> str:
    return (
        s.replace("\\n", "\n")
        .replace("\\r", "")
        .replace('\\"', '"')
        .replace("\\\\", "\\")
    )


def _escape_js(s: str) -> str:
    return (
        s.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", "\\n")
        .replace("\r", "")
    )


def _i18n_literal(zh_hans: str, zh_hant: str, en: str) -> str:
    return (
        '{ "zh-Hans": "'
        + _escape_js(zh_hans)
        + '", "zh-Hant": "'
        + _escape_js(zh_hant)
        + '", en: "'
        + _escape_js(en)
        + '" }'
    )


def export_xlsx(works: list[dict]) -> None:
    wb = Workbook()
    guide = wb.active
    guide.title = "使用说明"
    guide["A1"] = EDITABLE_NOTE
    guide["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    guide.column_dimensions["A"].width = 96
    guide.row_dimensions[1].height = 140

    ws = wb.create_sheet("作品资料", 0)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4A5560")
    edit_fill = PatternFill("solid", fgColor="FFF7E6")
    lock_fill = PatternFill("solid", fgColor="F0F0F0")

    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(1, col, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for idx, work in enumerate(works, 1):
        row = [
            idx,
            work["image_file"],
            work["titles"].get("zh-Hans", ""),
            work["year"],
            work["materials"].get("zh-Hans", ""),
            work["notes"].get("zh-Hans", ""),
            work["processCaptions"].get("zh-Hans", ""),
            work["titles"].get("en", ""),
            work["materials"].get("en", ""),
            work["notes"].get("en", ""),
            work["processCaptions"].get("en", ""),
        ]
        for col, value in enumerate(row, 1):
            cell = ws.cell(idx + 1, col, value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if col in (1, 2):
                cell.fill = lock_fill
            elif col in (3, 4, 5, 6, 7):
                cell.fill = edit_fill

    widths = [6, 42, 16, 8, 14, 48, 36, 22, 18, 48, 36]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(works) + 1}"

    wb.save(XLSX_PATH)
    print(f"已导出 {len(works)} 件作品到 {XLSX_PATH}")


def load_excel_rows() -> list[dict]:
    if not XLSX_PATH.exists():
        raise FileNotFoundError(f"找不到 {XLSX_PATH}")
    wb = load_workbook(XLSX_PATH)
    if "作品资料" not in wb.sheetnames:
        raise RuntimeError("Excel 中缺少工作表「作品资料」")
    ws = wb["作品资料"]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    expected = HEADERS
    if headers[: len(expected)] != expected:
        raise RuntimeError(f"表头不匹配。期望: {expected}，实际: {headers}")

    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or not values[1]:
            continue
        rows.append(
            {
                "序号": values[0],
                "图片文件": str(values[1]).strip(),
                "名称": (values[2] or "").strip() if values[2] is not None else "",
                "年份": str(values[3]).strip() if values[3] is not None else "",
                "创作媒介": (values[4] or "").strip() if values[4] is not None else "",
                "一句话介绍": (values[5] or "").strip() if values[5] is not None else "",
                "创作过程说明": (values[6] or "").strip() if values[6] is not None else "",
                "名称（英文）": (values[7] or "").strip() if len(values) > 7 and values[7] is not None else "",
                "创作媒介（英文）": (values[8] or "").strip() if len(values) > 8 and values[8] is not None else "",
                "一句话介绍（英文）": (values[9] or "").strip() if len(values) > 9 and values[9] is not None else "",
                "创作过程说明（英文）": (values[10] or "").strip() if len(values) > 10 and values[10] is not None else "",
            }
        )
    return rows


def rebuild_work_object(old: dict, row: dict) -> str:
    title_zh = row["名称"] or old["titles"].get("zh-Hans", "")
    material_zh = row["创作媒介"] or old["materials"].get("zh-Hans", "")
    note_zh = row["一句话介绍"] or old["notes"].get("zh-Hans", "")
    process_zh = row["创作过程说明"]

    title_en = row["名称（英文）"] or old["titles"].get("en", "")
    material_en = row["创作媒介（英文）"] or old["materials"].get("en", "")
    note_en = row["一句话介绍（英文）"] or old["notes"].get("en", "")
    process_en = row["创作过程说明（英文）"] or old["processCaptions"].get("en", "")

    year = row["年份"] or old["year"]
    image_file = row["图片文件"] or old["image_file"]

    lines = [
        "        {",
        f"          titles: {_i18n_literal(title_zh, s2t(title_zh), title_en)},",
        f'          year: "{_escape_js(year)}",',
        f"          materials: {_i18n_literal(material_zh, s2t(material_zh), material_en)},",
        f'          image: art("{_escape_js(image_file)}"),',
        f'          size: "{_escape_js(old["size"])}",',
    ]
    if old.get("series"):
        lines.append(f'          series: "{_escape_js(old["series"])}",')
    if old.get("panel"):
        lines.append(f'          panel: "{_escape_js(old["panel"])}",')
    if old.get("process_file"):
        lines.append(f'          processImage: art("{_escape_js(old["process_file"])}"),')
        # Prefer Excel process caption; fall back to previous zh-Hans then convert.
        process_src = process_zh or old["processCaptions"].get("zh-Hans", "")
        if process_src or process_en or old["processCaptions"]:
            lines.append("          processCaptions: {")
            lines.append(
                f'            "zh-Hans": "{_escape_js(process_src)}",'
            )
            lines.append(
                f'            "zh-Hant": "{_escape_js(s2t(process_src))}",'
            )
            lines.append(f'            en: "{_escape_js(process_en)}"')
            lines.append("          },")
    lines.append(f"          notes: {_i18n_literal(note_zh, s2t(note_zh), note_en)}")
    lines.append("        }")
    return "\n".join(lines)


def import_xlsx() -> None:
    html = HTML_PATH.read_text(encoding="utf-8")
    works = parse_works(html)
    by_file = {w["image_file"]: w for w in works}
    rows = load_excel_rows()

    if len(rows) != len(works):
        print(
            f"警告：Excel 有 {len(rows)} 行，网页有 {len(works)} 件作品。将按图片文件名匹配。",
            file=sys.stderr,
        )

    rebuilt = []
    seen = set()
    for row in rows:
        key = row["图片文件"]
        if key not in by_file:
            raise RuntimeError(f"Excel 中的图片文件无法匹配网页作品: {key}")
        rebuilt.append(rebuild_work_object(by_file[key], row))
        seen.add(key)

    missing = [w["image_file"] for w in works if w["image_file"] not in seen]
    if missing:
        raise RuntimeError("Excel 缺少这些作品行: " + ", ".join(missing))

    prefix, _, suffix = extract_works_block(html)
    new_block = "const works = [\n" + ",\n".join(rebuilt) + "\n      ];"
    HTML_PATH.write_text(prefix + new_block + suffix, encoding="utf-8")
    print(f"已从 {XLSX_PATH} 同步 {len(rebuilt)} 件作品到 {HTML_PATH}")
    print("简体内容已写入；繁体已按简体自动转换；英文优先使用 Excel 英文列，否则保留原英文。")


def main() -> None:
    parser = argparse.ArgumentParser(description="Sync gallery works between Excel and index.html")
    parser.add_argument("command", choices=["export", "import"], help="export=网页导出Excel；import=Excel写回网页")
    args = parser.parse_args()

    html = HTML_PATH.read_text(encoding="utf-8")
    works = parse_works(html)
    if args.command == "export":
        export_xlsx(works)
    else:
        import_xlsx()
        # Round-trip sanity: parse again
        parse_works(HTML_PATH.read_text(encoding="utf-8"))
        print("二次解析通过。")


if __name__ == "__main__":
    main()
