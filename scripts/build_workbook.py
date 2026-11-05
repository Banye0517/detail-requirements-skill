"""Export a reviewed detail-request JSON to an image-embedded Excel workbook."""

import argparse
import json
import os
from pathlib import Path
import tempfile
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image as PillowImage


HEADER = "24556A"
PALE = "EFF6F9"
INPUT = "FFF2CF"
WHITE = "FFFFFF"


def safe_text(value):
    text = str(value or "")
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def valid_url(value):
    parsed = urlparse(value or "")
    return parsed.scheme in ("https", "http") and bool(parsed.hostname)


def set_cell(sheet, row, column, value, fill=None, bold=False):
    cell = sheet.cell(row, column, safe_text(value))
    cell.font = Font(name="Arial", size=11, bold=bold, color="183C4A")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    return cell


def band(sheet, row, text, columns=9):
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=columns)
    cell = set_cell(sheet, row, 1, text, HEADER, True)
    cell.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    for col in range(1, columns + 1):
        sheet.cell(row, col).fill = PatternFill("solid", fgColor=HEADER)
    sheet.row_dimensions[row].height = 30


def write_records(workbook, name, headers, records):
    sheet = workbook.create_sheet(name)
    for col, title in enumerate(headers, 1):
        set_cell(sheet, 1, col, title, HEADER, True).font = Font(
            name="Arial", size=11, bold=True, color=WHITE
        )
        sheet.column_dimensions[get_column_letter(col)].width = 40 if col > 1 else 22
    sheet.row_dimensions[1].height = 28
    for row, values in enumerate(records, 2):
        for col, value in enumerate(values, 1):
            cell = set_cell(sheet, row, col, value, PALE if row % 2 == 0 else WHITE)
            if valid_url(str(value)):
                cell.hyperlink = str(value)
                cell.font = Font(name="Arial", size=11, color="0563C1", underline="single")
        sheet.row_dimensions[row].height = 80
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    return sheet


def template_data():
    tasks = ["产品识别", "购买理由", "卖点强化", "卖点强化", "卖点强化", "使用场景", "产品信息"]
    return {
        "product_name": "产品名称待填写",
        "selling_points": ["填写用户提供的原始卖点"],
        "audience": "待确认",
        "summary": "智能体调研后填写文案、执行方案和三张详情参考图。",
        "screens": [{"id": f"第{i}屏", "task": task, "references": []} for i, task in enumerate(tasks, 1)],
        "main_images": [{"id": f"主图{i}", "task": "白底图" if i == 5 else "产品识别或购买理由", "references": []} for i in range(1, 6)],
        "research": [],
        "brand_search": [],
        "pending": [],
    }


def export(data, output, base, allow_incomplete=False, blank=False):
    if not str(data.get("product_name", "")).strip():
        raise ValueError("product_name is required")
    screens = data.get("screens", [])
    if not screens:
        raise ValueError("At least one detail screen is required")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "详情需求表"
    widths = [14, 19, 25, 33, 47, 38, 31, 31, 31]
    for col, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    band(sheet, 1, data["product_name"] + " · 详情需求表")
    selling_points = data.get("selling_points", [])
    if isinstance(selling_points, list):
        selling_points = "；".join(map(str, selling_points))
    overview = [
        ("原始卖点", selling_points),
        ("目标人群", data.get("audience", "待确认")),
        ("交付规格", "详情单屏9:16，900×1600 px；每屏3张参考图，图下逐张链接。主图尺寸按渠道。"),
        ("草稿状态", data.get("summary", "运营审核草稿，待补项见右列及运营待补表。")),
    ]
    for row, (label, value) in enumerate(overview, 2):
        set_cell(sheet, row, 1, label, PALE, True)
        sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        set_cell(sheet, row, 2, value, PALE)
        sheet.row_dimensions[row].height = 38
    headers = ["屏号", "本屏任务", "主标题", "副标题 / 标签 / 小字", "画面执行与设计理由", "运营待确认", "详情参考图1", "详情参考图2", "详情参考图3"]
    for col, title in enumerate(headers, 1):
        set_cell(sheet, 7, col, title, HEADER, True).font = Font(name="Arial", size=11, bold=True, color=WHITE)
    sheet.row_dimensions[7].height = 28
    row = 8
    image_count = 0
    link_count = 0
    gaps = []
    groups = [("详情页", screens), ("主图建议", data.get("main_images", []))]
    for group_name, blocks in groups:
        if not blocks:
            continue
        band(sheet, row, group_name)
        row += 1
        for index, block in enumerate(blocks):
            references = block.get("references", [])
            if len(references) > 3:
                raise ValueError(f"{block.get('id')} has more than three references")
            if not (allow_incomplete or blank) and len(references) != 3:
                raise ValueError(f"{block.get('id')} requires exactly three references")
            headline = block.get("headline", "")
            subtitle = block.get("subtitle", "")
            if group_name == "详情页" and not blank:
                if not 6 <= len(headline) <= 12:
                    raise ValueError(f"Headline must be 6-12 characters: {headline}")
                if not 8 <= len(subtitle) <= 16:
                    raise ValueError(f"Subtitle must be 8-16 characters: {subtitle}")
            copy_parts = [subtitle, block.get("labels", ""), block.get("tips", "")]
            values = [block.get("id", ""), block.get("task", ""), headline,
                      "\n".join(str(x) for x in copy_parts if x),
                      block.get("visual", ""), block.get("pending", "")]
            for col, value in enumerate(values, 1):
                sheet.merge_cells(start_row=row, start_column=col, end_row=row + 2, end_column=col)
                set_cell(sheet, row, col, value, INPUT if col == 6 else (PALE if index % 2 == 0 else WHITE))
            sheet.row_dimensions[row].height = 252
            sheet.row_dimensions[row + 1].height = 48
            sheet.row_dimensions[row + 2].height = 72
            for slot in range(3):
                col = slot + 7
                if slot >= len(references):
                    gaps.append(f"{block.get('id')} 图{slot + 1}")
                    set_cell(sheet, row, col, "9:16详情参考图待补", INPUT)
                    set_cell(sheet, row + 1, col, "商品来源链接待补", INPUT)
                    continue
                ref = references[slot]
                url = ref.get("product_url", "")
                image_path = Path(ref.get("image", ""))
                if not image_path.is_absolute():
                    image_path = base / image_path
                if not valid_url(url) or not ref.get("brand") or not ref.get("product_name"):
                    raise ValueError(f"Incomplete source metadata at {block.get('id')} slot {slot + 1}")
                if not image_path.is_file():
                    raise ValueError(f"Missing reference image: {image_path}")
                with PillowImage.open(image_path) as bitmap:
                    width, height = bitmap.size
                    if abs(width / height - 9 / 16) > 0.003:
                        raise ValueError(f"Reference must be 9:16: {image_path}")
                    bitmap.verify()
                image = Image(str(image_path))
                image.width = 180
                image.height = 320
                sheet.add_image(image, f"{get_column_letter(col)}{row}")
                image_count += 1
                source = set_cell(sheet, row + 1, col, f"{ref['brand']} · {ref['product_name']}\n打开商品详情")
                source.hyperlink = url
                source.font = Font(name="Arial", size=10, color="0563C1", underline="single")
                link_count += 1
                note = f"借鉴：{ref.get('borrow', '')}\n本品：{ref.get('application', '')}"
                if ref.get("crop_note"):
                    note += "\n" + ref["crop_note"]
                set_cell(sheet, row + 2, col, note)
            row += 3
    sheet.freeze_panes = "G8"
    sheet.sheet_view.showGridLines = False
    sheet.print_options.horizontalCentered = True
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:7"
    sheet.print_area = f"A1:I{row - 1}"
    research_rows = [[x.get(k, "") for k in ("brand", "product_name", "product_url", "observed", "application", "limits")] for x in data.get("research", [])]
    research = write_records(workbook, "竞品调研", ["品牌", "商品", "商品直链", "实际观察", "本品应用", "证据范围与限制"], research_rows)
    start = len(research_rows) + 3
    for offset, record in enumerate(data.get("brand_search", [])):
        for col, key in enumerate(("brand", "status", "source"), 1):
            cell = set_cell(research, start + offset, col, record.get(key, ""), PALE)
            if key == "source" and valid_url(record.get(key, "")):
                cell.hyperlink = record[key]
        research.row_dimensions[start + offset].height = 60
    pending_rows = [[x.get(k, "") for k in ("owner", "item", "required", "reason")] for x in data.get("pending", [])]
    write_records(workbook, "运营待补", ["责任方", "事项", "所需资料", "影响范围"], pending_rows)
    if gaps and not blank:
        set_cell(sheet, 5, 2, "参考图未完整，缺口：" + "、".join(gaps), INPUT)
    output.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=output.parent)
    os.close(handle)
    temp = Path(temp_name)
    try:
        workbook.save(temp)
        reopened = load_workbook(temp)
        check = reopened["详情需求表"]
        actual_links = sum(bool(cell.hyperlink) for cells in check for cell in cells)
        if len(check._images) != image_count or actual_links != link_count:
            raise ValueError("Image or hyperlink export verification failed")
        reopened.close()
        os.replace(temp, output)
    finally:
        if temp.exists():
            temp.unlink()
    print(json.dumps({"output": str(output), "images": image_count, "links": link_count, "reference_gaps": gaps if not blank else []}, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="UTF-8 JSON input, or blank-template")
    parser.add_argument("output", type=Path)
    parser.add_argument("--allow-incomplete", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    if args.output.exists() and not args.overwrite:
        parser.error("Output already exists; choose a new name or use --overwrite")
    blank = args.input == "blank-template"
    path = Path(args.input)
    data = template_data() if blank else json.loads(path.read_text(encoding="utf-8"))
    export(data, args.output, path.resolve().parent, args.allow_incomplete, blank)


if __name__ == "__main__":
    main()
