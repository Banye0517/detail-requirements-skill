"""复制工单并新增 J 列「本屏线框」：图片锚在每屏第 1 行，尺寸同 G-I；不改原件。"""
import copy, json, shutil, sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.worksheet.cell_range import CellRange

HERE = Path(__file__).resolve().parent


def resolve_source(argv):
    """Prefer an explicit path, then a local source-copy, then the official example."""
    if len(argv) > 1:
        return Path(argv[1])
    local = HERE / "source-copy.xlsx"
    if local.is_file():
        return local
    sibling = HERE.parent / "angus-beef" / "detail-request.xlsx"
    if sibling.is_file():
        return sibling
    return local


SRC = resolve_source(sys.argv)
OUT = HERE / "detail-request-with-wireframe.xlsx"

wb = load_workbook(SRC)
ws = wb["详情需求表"]
before_imgs = len(ws._images)
before_links = sum(bool(c.hyperlink) for row in ws for c in row)

ws.column_dimensions["J"].width = ws.column_dimensions["I"].width  # 31

# 跨 A-I 的合并区扩到 J
for rng in list(ws.merged_cells.ranges):
    if rng.max_col == 9 and rng.min_row == rng.max_row:
        r, c0 = rng.min_row, rng.min_col
        ws.unmerge_cells(str(rng))
        src = ws.cell(r, 9)
        ws.cell(r, 10)._style = copy.copy(src._style)
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=10)

ws["J7"] = "本屏线框"
ws["J7"]._style = copy.copy(ws["I7"]._style)
ws["B4"] = (
    str(ws["B4"].value)
    + " J列「本屏线框」为按本屏工单文案生成的黑白结构线稿（试看），不是详情参考图。"
    + "主图线框按 900×1600（9:16）截出，只为与表格参考图列对齐，不是已确认的渠道比例。"
)

manifest = json.loads((HERE / "manifest.json").read_text("utf-8"))
shots = sorted((HERE / "screens").glob("*.png"))
assert len(shots) == len(manifest) == 12, (len(shots), len(manifest))
ref = ws._images[0]
for m, png in zip(manifest, shots):
    r = m["row"]
    img = Image(str(png))
    img.width, img.height = 180, 320           # 与 G-I 显示尺寸一致（ext 1714500x3048000 EMU）
    ws.add_image(img, f"J{r}")
    for k in range(3):
        ws.cell(r + k, 10)._style = copy.copy(ws.cell(r + k, 9)._style)
    ws.cell(r, 10).value = None
    c = ws.cell(r + 1, 10, "本屏黑白线框")
    c.hyperlink = None
    note = f"按本屏工单文案自动生成的结构线稿，只示意层级与图片位，不是真实详情参考。模板：{m['template']}"
    if "主图" in str(m.get("label", "")) or m.get("template") == "主图组成框":
        note += "。主图线框按 900×1600（9:16）只为与参考图列对齐，真实渠道比例待确认。"
        ws.row_dimensions[r + 2].height = 96
    ws.cell(r + 2, 10, note)
    # 链接行字体去掉下划线/蓝色
    from openpyxl.styles import Font
    ws.cell(r + 1, 10).font = Font(name="Arial", size=10, color="183C4A", bold=True)

ws.print_area = f"A1:J{ws.max_row}"
wb.save(OUT)

chk = load_workbook(OUT)["详情需求表"]
imgs = chk._images
links = sum(bool(c.hyperlink) for row in chk for c in row)
j = [i for i in imgs if i.anchor._from.col == 9]
print(json.dumps(dict(before_imgs=before_imgs, after_imgs=len(imgs), j_imgs=len(j),
    j_rows=[i.anchor._from.row + 1 for i in j], ext=[(i.anchor.ext.cx, i.anchor.ext.cy) for i in j][:1],
    gi_ext=list({(i.anchor.ext.cx, i.anchor.ext.cy) for i in imgs if i.anchor._from.col < 9}),
    links_before=before_links, links_after=links, merges=len(chk.merged_cells.ranges),
    widthJ=chk.column_dimensions["J"].width), ensure_ascii=False))
