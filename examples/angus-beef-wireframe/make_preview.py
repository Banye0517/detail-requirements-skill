"""预览拼图：前 4 屏详情，每行 = 屏号/文案 + 3 张参考图（取自工单内嵌图）+ 右侧本屏线框。"""
import io, json
from pathlib import Path
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent


def font_file(bold=False):
    """Find a system CJK font. This repo does not bundle font files."""
    preferred = [
        "NotoSansCJK-Bold.ttc" if bold else "NotoSansCJK-Regular.ttc",
        "NotoSansCJKsc-Bold.otf" if bold else "NotoSansCJKsc-Regular.otf",
        "NotoSansSC-Bold.otf" if bold else "NotoSansSC-Regular.otf",
    ]
    fallback = ["wqy-microhei.ttc", "DroidSansFallbackFull.ttf"]
    roots = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".local/share/fonts"]
    by_name = {}
    for root in roots:
        if not root.is_dir():
            continue
        for path in root.rglob("*"):
            if path.is_file() and path.name not in by_name:
                by_name[path.name] = path
    for name in preferred + fallback:
        if name in by_name:
            return by_name[name]
    raise SystemExit("重做预览需要本机的中文字体（优先 Noto Sans CJK）。本仓库不附带字体文件。")


def font(sz, bold=False):
    path = font_file(bold)
    index = 2 if path.suffix.lower() == ".ttc" and "NotoSansCJK" in path.name else 0
    return ImageFont.truetype(str(path), sz, index=index)
HEAD, PALE, INPUT, INK, LINK = (36, 85, 106), (239, 246, 249), (255, 242, 207), (24, 60, 74), (5, 99, 193)

ws = load_workbook(HERE / "detail-request-with-wireframe.xlsx")["详情需求表"]
imgs = {(i.anchor._from.col, i.anchor._from.row + 1): i for i in ws._images}
manifest = json.loads((HERE / "manifest.json").read_text("utf-8"))[:4]

LW, IW, IH, GAP, PAD = 300, 240, 427, 16, 24
W = PAD * 2 + LW + 4 * (IW + GAP)
ROWH = IH + 110
H = 150 + 44 + len(manifest) * ROWH + 50
im = Image.new("RGB", (W, H), "white"); d = ImageDraw.Draw(im)

def wrap(text, f, width):
    out, cur = [], ""
    for ch in text:
        if ch == "\n": out.append(cur); cur = ""; continue
        if d.textlength(cur + ch, font=f) > width: out.append(cur); cur = ch
        else: cur += ch
    if cur: out.append(cur)
    return out

def para(x, y, text, f, width, fill=INK, gap=6, maxl=99):
    for ln in wrap(text, f, width)[:maxl]:
        d.text((x, y), ln, font=f, fill=fill); y += f.size + gap
    return y

d.rectangle([0, 0, W, 64], fill=HEAD)
d.text((PAD, 14), "安格斯牛肉脆 · 详情需求表（示例）· 新增 J 列本屏线框", font=font(26, True), fill="white")
para(PAD, 80, "每屏横排 3 张真实详情参考图（9:16），右侧 J 列为按本屏工单文案自动生成的黑白结构线框，只示意分屏、文字层级和图片位置，"
     "不是真实详情参考。截取前 4 屏；完整工单 7 屏详情 + 5 张主图，共 12 张线框。", font(15), W - 2 * PAD, (60, 80, 90))
y = 150
d.rectangle([0, y, W, y + 44], fill=HEAD)
cols = ["屏号 / 任务 / 文案", "详情参考图1", "详情参考图2", "详情参考图3", "J · 本屏线框（新增）"]
xs = [PAD, PAD + LW] + [PAD + LW + k * (IW + GAP) for k in (1, 2, 3)]
for x, t in zip(xs, cols): d.text((x, y + 10), t, font=font(16, True), fill="white")
y += 44
for n, m in enumerate(manifest):
    r = m["row"]
    bg = PALE if n % 2 == 0 else (255, 255, 255)
    d.rectangle([0, y, W, y + ROWH], fill=bg)
    ty = y + 18
    d.text((PAD, ty), str(ws.cell(r, 1).value), font=font(16, True), fill=HEAD); ty += 28
    d.text((PAD, ty), "任务：" + str(ws.cell(r, 2).value), font=font(14), fill=INK); ty += 30
    ty = para(PAD, ty, str(ws.cell(r, 3).value), font(22, True), LW - 30, INK, 6) + 4
    ty = para(PAD, ty, str(ws.cell(r, 4).value), font(15), LW - 30, INK, 5, 4) + 8
    ty = para(PAD, ty, "画面：" + str(ws.cell(r, 5).value), font(13), LW - 30, (90, 100, 110), 4, 5) + 8
    lines = wrap("待确认：" + str(ws.cell(r, 6).value), font(13), LW - 46)[:4]
    d.rectangle([PAD - 6, ty - 4, PAD + LW - 30, ty + len(lines) * 19 + 6], fill=INPUT)
    for ln in lines: d.text((PAD, ty), ln, font=font(13), fill=INK); ty += 19
    for k in range(4):
        x = PAD + LW + k * (IW + GAP)
        pic = Image.open(io.BytesIO(imgs[(6 + k, r)]._data())).convert("RGB").resize((IW, IH), Image.LANCZOS)
        im.paste(pic, (x, y + 18))
        if k == 3:
            d.rectangle([x - 1, y + 17, x + IW, y + 18 + IH], outline=(17, 17, 17), width=2)
        ny = y + 18 + IH + 8
        cell = ws.cell(r + 1, 7 + k).value or ""
        first = str(cell).split("\n")[0]
        para(x, ny, first, font(12), IW, LINK if k < 3 else INK, 3, 2)
        note = str(ws.cell(r + 2, 7 + k).value or "").split("\n")[0]
        para(x, ny + 36, note, font(11), IW, (90, 100, 110), 3, 3)
    y += ROWH
d.text((PAD, y + 14), "第三方商品图片归原权利人，仅作调研与设计参考展示。线框由脚本按工单文案生成，「待补」为运营待提供项。", font=font(13), fill=(120, 120, 120))
im.save(HERE / "preview-with-wireframe.jpg", quality=92)
print(im.size)
