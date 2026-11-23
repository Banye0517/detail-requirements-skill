"""从详情需求表 xlsx 读取每屏文案，按固定模板生成黑白线框 HTML（交付版 + 截图版）。
文字只取自工单单元格；模板只决定版式，不新增产品信息。"""
import html, json, re, sys
from pathlib import Path
from openpyxl import load_workbook

def resolve_source(argv):
    """Prefer an explicit path, then a local source-copy, then the official example."""
    here = Path(__file__).resolve().parent
    if len(argv) > 1:
        return Path(argv[1])
    local = here / "source-copy.xlsx"
    if local.is_file():
        return local
    sibling = here.parent / "angus-beef" / "detail-request.xlsx"
    if sibling.is_file():
        return sibling
    return local


SRC = resolve_source(sys.argv)
OUT = Path(__file__).resolve().parent
E = html.escape


def read_blocks(path):
    ws = load_workbook(path)["详情需求表"]
    blocks, group = [], None
    for r in range(8, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a in ("详情页", "主图建议"):
            group = a
            continue
        if a and ws.row_dimensions[r].height == 252:
            v = [str(ws.cell(r, c).value or "") for c in range(1, 7)]
            blocks.append(dict(group=group, row=r, id=v[0], task=v[1], headline=v[2],
                               copy=v[3], visual=v[4], pending=v[5]))
    product = str(ws["A1"].value).split(" · ")[0]
    return product, blocks


def lines(text):
    return [x.strip() for x in text.split("\n") if x.strip()]


def field(text, key):
    """取『key：值』行的值。"""
    for ln in lines(text):
        m = re.match(rf"^{key}(（[^）]*）)?[：:](.+)$", ln)
        if m:
            return m.group(2).strip()
    return ""


def split_items(value):
    return [x.strip() for x in re.split(r"[｜|／/]", value) if x.strip()]


def sentences(visual):
    """画面执行拆句，去掉『9:16』规格句。"""
    return [s.strip() for s in re.split(r"[。；]", visual) if s.strip() and s.strip() != "9:16"]


def screen_label(b):
    m = re.search(r"第(\d+)屏", b["id"])
    if m:
        return f"详情 · 第 {m.group(1)} 屏"
    m = re.search(r"主图(\d+)", b["id"])
    return f"主图 · {m.group(1)}" if m else b["id"]


def subtitle(b):
    return lines(b["copy"])[0] if lines(b["copy"]) else ""


def box(label, note="", cls=""):
    n = f'<span class="note">{E(note)}</span>' if note else ""
    return f'<div class="ph {cls}"><span class="tag">{E(label)}</span>{n}</div>'


def head(b, extra=""):
    return (f'<div class="meta"><span>{E(screen_label(b))}</span><span>{E(b["task"])}{extra}</span></div>'
            f'<h1>{E(b["headline"])}</h1><p class="sub">{E(subtitle(b))}</p>')


# ---------- 模板 ----------
def t_hero(b):
    s = sentences(b["visual"])
    icons = split_items(field(b["copy"], "icon"))
    name = field(b["copy"], "品名")
    dots = "".join(f'<div class="dot"><i></i><b>{E(x)}</b></div>' for x in icons)
    return (head(b) + box("主视觉图位", s[0] if s else "", "grow") +
            f'<div class="dots">{dots}</div>' +
            (f'<p class="foot">品名：{E(name)}</p>' if name else ""))


def t_material(b):
    s = sentences(b["visual"])
    return (head(b) + box("原料特写图位", s[0] if s else "", "grow") +
            '<div class="row2">' + box("产地示意位", s[1] if len(s) > 1 else "", "small") +
            box("原料证明位", "待补", "small dashed") + '</div>')


def t_ingredients(b):
    s = sentences(b["visual"])
    m = re.search(r"只有(.+?)和(.+?)$", b["headline"])
    items = [m.group(1), m.group(2)] if m else ["配料1", "配料2"]
    cards = "".join(f'<div class="card"><div class="circle"></div><b>{E(x)}</b><span>配料卡</span></div>' for x in items)
    plus = '<div class="plus">+</div>'
    return (head(b) + f'<div class="cards">{cards.replace("</div><div class=\"card\">", "</div>" + plus + "<div class=\"card\">", 1)}</div>' +
            box("真实背标位", "配料表截图（待补）", "mid dashed") +
            f'<p class="foot">{E(s[0]) if s else ""}</p>')


def t_process(b):
    s = sentences(b["visual"])
    flow = next((x for x in s if "→" in x), "")
    nodes = [x.strip() for x in re.sub(r"用.*?节点连接", "", flow).split("→") if x.strip()]
    nodes = [re.sub(r"^.*?(原料)$", r"\1", n) for n in nodes]
    html_nodes = '<div class="line"></div>'.join(
        f'<div class="node"><div class="circle"><em>{i}</em></div><b>{E(n)}</b></div>' for i, n in enumerate(nodes, 1))
    return (head(b) + box("烘盘 / 成品图位", s[0] if s else "", "grow") +
            f'<div class="flow">{html_nodes}</div>' +
            '<p class="foot">温度、时长等参数：待补</p>')


def t_texture(b):
    s = sentences(b["visual"])
    return (head(b) + '<div class="macro"><div class="ring"><span class="tag">近景图位</span>'
            f'<span class="note">{E(s[0]) if s else ""}</span></div></div>' +
            box("成品堆叠位", s[1] if len(s) > 1 else "", "strip") +
            '<p class="foot">口感描述：待试吃反馈后补</p>')


def t_scenes(b):
    labels = split_items(field(b["copy"], "场景标签"))
    shorts = split_items(field(b["copy"], "场景短句"))
    cells = ""
    for i, lab in enumerate(labels):
        sh = shorts[i] if i < len(shorts) else ""
        cells += (f'<div class="cell"><div class="ph in"><span class="tag">场景图位</span></div>'
                  f'<b>{E(lab)}</b><span>{E(sh)}</span></div>')
    return head(b, ' · <em class="prop">提案</em>') + f'<div class="grid">{cells}</div>'


def t_info(b):
    vis, pend = b["visual"], b["pending"]
    m = re.search(r"。([^。]*?)逐项列出", vis)
    rows = [x for x in re.split(r"[、，]", m.group(1)) if x] if m else []
    for extra in re.split(r"[、，与]", pend.split("均待")[0]):
        extra = extra.replace("成品", "").strip()
        if extra and extra not in rows and extra not in ("规格",):
            rows.append(extra)
    trs = ""
    for r in rows:
        if r == "产品名" and b.get("product"):
            trs += f'<tr><th>{E(r)}</th><td>{E(b["product"])}</td></tr>'
        else:
            trs += f'<tr><th>{E(r)}</th><td class="tbd">待补</td></tr>'
    return (head(b) + '<div class="row2">' + box("包装正面图位", "待补", "pack") +
            box("包装背面图位", "待补", "pack") + '</div>' +
            f'<table>{trs}</table>')


def t_main(b):
    s = sentences(b["visual"])
    t = b["task"]
    if "原料" in t:
        body = box("原料近景图位", s[0] if s else "", "grow") + box("来源标记位", "", "badge")
    elif "配料" in t:
        body = ('<div class="row2 grow">' + box("配料分区 1", "", "") + box("配料分区 2", "", "") + '</div>' +
                box("背标证据位", s[0] if s else "", "mid dashed"))
    elif "工艺" in t:
        body = box("烘盘图位", s[0] if s else "", "half") + box("成品图位", "", "half")
    elif "白底" in t:
        body = '<div class="white grow">' + box("包装正面实物位", s[0] if s else "", "pack2") + '</div>'
    else:
        body = '<div class="row2 grow">' + box("包装图位", "", "") + box("成品特写位", "", "") + '</div>' + \
               (f'<p class="foot">{E(s[0])}</p>' if s else "")
    return head(b) + body


def pick(b):
    if b["group"] == "主图建议":
        return t_main, "主图组成框"
    t = b["task"]
    for key, fn, name in [("是什么", t_hero, "首屏识别"), ("原料", t_material, "原料大图"),
                          ("配料", t_ingredients, "配料卡"), ("工艺", t_process, "三节点流程"),
                          ("质地", t_texture, "近景微距"), ("场景", t_scenes, "四宫格场景"),
                          ("规格", t_info, "信息表")]:
        if key in t:
            return fn, name
    return t_main, "通用组成框"


CSS = """
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#fff;color:#111;font-family:%FONT%;-webkit-font-smoothing:antialiased}
.page{width:375px;margin:0 auto;background:#fff}
.cover{height:120px;padding:22px 20px;border-bottom:1px solid #111}
.cover h2{font-size:18px;font-weight:800}.cover p{font-size:11px;color:#666;margin-top:6px;line-height:1.5}
.band{height:40px;line-height:40px;padding:0 20px;font-size:12px;font-weight:700;letter-spacing:2px;background:#111;color:#fff}
.screen{width:375px;height:667px;padding:16px 22px 18px;border-bottom:1px solid #111;display:flex;flex-direction:column;gap:12px;overflow:hidden;background:#fff}
.meta{display:flex;justify-content:space-between;font-size:10px;color:#888;letter-spacing:.5px}
.prop{font-style:normal;border:1px solid #111;color:#111;padding:0 4px;border-radius:2px}
h1{font-size:30px;line-height:1.2;font-weight:900;letter-spacing:.5px}
.sub{font-size:15px;color:#444;margin-top:-4px}
.ph{background:#eee;border:1px solid #cfcfcf;border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;padding:10px;text-align:center;position:relative}
.ph::before,.ph::after{content:"";position:absolute;inset:0;background:linear-gradient(to top right,transparent calc(50% - .5px),#d4d4d4,transparent calc(50% + .5px));pointer-events:none;border-radius:8px}
.ph::after{background:linear-gradient(to top left,transparent calc(50% - .5px),#d4d4d4,transparent calc(50% + .5px))}
.ph .tag{position:relative;z-index:1;font-size:12px;font-weight:700;background:#fff;border:1px solid #111;padding:2px 8px;border-radius:10px}
.ph .note{position:relative;z-index:1;font-size:11px;color:#555;background:#eee;padding:2px 6px;line-height:1.5;max-width:90%}
.dashed{border:1.5px dashed #999;background:#f7f7f7}.dashed::before,.dashed::after{display:none}
.grow{flex:1;min-height:0}.small{flex:1;height:96px}.mid{height:110px}.strip{height:92px}.pack{flex:1;height:120px}
.half{flex:1}.badge{height:64px;width:160px;align-self:flex-start;border-radius:32px}
.row2{display:flex;gap:12px}.row2.grow>.ph{flex:1}
.white{border:1px solid #111;border-radius:8px;padding:20px;display:flex}.white .ph{flex:1}
.dots{display:flex;justify-content:space-between;padding:0 6px}
.dot{display:flex;flex-direction:column;align-items:center;gap:6px;width:90px}
.dot i{width:58px;height:58px;border-radius:50%;border:1.5px solid #111;background:#eee}
.dot b{font-size:13px;font-weight:700}
.foot{font-size:11px;color:#666;line-height:1.5}
.cards{display:flex;align-items:center;gap:10px;flex:1}
.card{flex:1;height:100%;max-height:300px;background:#f2f2f2;border:1px solid #cfcfcf;border-radius:10px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px}
.card .circle{width:92px;height:92px;border-radius:50%;border:1.5px solid #111;background:#fff}
.card b{font-size:22px;font-weight:900}.card span{font-size:11px;color:#888}
.plus{font-size:26px;font-weight:300}
.flow{display:flex;align-items:flex-start;padding:4px 2px 0}
.node{display:flex;flex-direction:column;align-items:center;gap:8px;width:84px}
.node .circle{width:64px;height:64px;border-radius:50%;border:1.5px solid #111;background:#eee;display:flex;align-items:center;justify-content:center}
.node em{font-style:normal;font-size:20px;font-weight:800}
.node b{font-size:13px;font-weight:700;white-space:nowrap}
.line{flex:1;height:1.5px;background:#111;margin-top:32px;position:relative}
.line::after{content:"";position:absolute;right:-1px;top:-4px;border-left:7px solid #111;border-top:5px solid transparent;border-bottom:5px solid transparent}
.macro{flex:1;display:flex;align-items:center;justify-content:center;min-height:0}
.ring{width:290px;height:290px;border-radius:50%;border:1.5px solid #111;background:#eee;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;padding:30px;text-align:center}
.ring .tag{font-size:12px;font-weight:700;background:#fff;border:1px solid #111;padding:2px 8px;border-radius:10px}
.ring .note{font-size:11px;color:#555;line-height:1.5}
.grid{flex:1;display:grid;grid-template-columns:1fr 1fr;gap:10px;min-height:0}
.cell{display:flex;flex-direction:column;gap:4px;min-height:0}
.cell .ph{flex:1}.cell b{font-size:15px;font-weight:800;margin-top:2px}.cell span{font-size:11px;color:#666}
table{width:100%;border-collapse:collapse;font-size:12px;flex:1}
th,td{border-top:1px solid #111;padding:0 8px;text-align:left;height:30px}
tr:last-child th,tr:last-child td{border-bottom:1px solid #111}
th{width:38%;font-weight:700;background:#f2f2f2}td.tbd{color:#999;letter-spacing:1px}
"""
FONT_DELIVER = '"PingFang SC","Hiragino Sans GB","Microsoft YaHei","Noto Sans CJK SC",sans-serif'
# 截图页只点名本机已安装字体，不把字体文件或本机 file:// 路径写进 HTML。
FACE = ('@font-face{font-family:"WFSans";font-weight:400;src:local("Noto Sans CJK SC")}'
        '@font-face{font-family:"WFSans";font-weight:700 900;src:local("Noto Sans CJK SC Bold")}')


def doc(title, body, shot=False):
    css = CSS.replace("%FONT%", '"WFSans","Noto Sans CJK SC",sans-serif' if shot else FONT_DELIVER)
    if shot:
        css = FACE + css
    return (f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=375,initial-scale=1"><title>{E(title)}</title>'
            f'<style>{css}</style></head><body>{body}</body></html>')


def main():
    product, blocks = read_blocks(SRC)
    sections, manifest, group = [], [], None
    for i, b in enumerate(blocks, 1):
        b["product"] = product
        fn, tname = pick(b)
        sid = f"s{i:02d}"
        sec = f'<section class="screen" id="{sid}">{fn(b)}</section>'
        if b["group"] != group:
            group = b["group"]
            sections.append(f'<div class="band">{E(group)}</div>')
        sections.append(sec)
        manifest.append(dict(id=sid, row=b["row"], label=b["id"], task=b["task"], headline=b["headline"],
                             template=tname, html=sec))
    cover = (f'<div class="cover"><h2>{E(product)} · 详情页黑白线框</h2>'
             '<p>按《详情需求表》每屏文案自动生成的结构线稿，只表达分屏、文字层级与图片位置；'
             '不是真实设计，也不是详情参考。灰框为图片位，「待补」为运营待提供。</p></div>')
    title = f"{product} · 详情页黑白线框"
    (OUT / "angus-wireframe.html").write_text(doc(title, f'<div class="page">{cover}{"".join(sections)}</div>'), "utf-8")
    (OUT / "_shot_full.html").write_text(doc(title, f'<div class="page">{cover}{"".join(sections)}</div>', True), "utf-8")
    sd = OUT / "_shots"; sd.mkdir(exist_ok=True)
    for m in manifest:
        (sd / f'{m["id"]}.html').write_text(doc(title, m.pop("html"), True), "utf-8")
    (OUT / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1), "utf-8")
    print(json.dumps([(m["id"], m["label"], m["template"]) for m in manifest], ensure_ascii=False))


if __name__ == "__main__":
    main()
