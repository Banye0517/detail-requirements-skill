# 安格斯牛肉脆 · 黑白线框试看

本目录是 `examples/angus-beef/detail-request.xlsx` 的试看稿。正式工单仍是那份没有 J 列的三图 Excel。Skill 日常仍按产品名称和卖点调研，并交付 Excel 需求表。打开本目录不会把主流程改成必须出线框。

## 试看怎么来的

`build_wireframe.py` 只读取工单单元格里的标题、副标题、标签和画面说明。它按本屏任务选用一个固定模板：首屏识别、原料大图、配料卡、三节点流程、近景微距、四宫格场景、信息表，主图则用主图组成框。模板只写「主视觉图位」这类位置名；工单里还没有的资料写成「待补」。

这是按这一份已有工单试做的示例，尚未在其他产品上验证，不能当作通用线框产品。

第 6 屏场景保持「提案」。第 7 屏产品信息除品名外保持「待补」。卖点原文和待核对项沿用工单，不在这里改写。普通零食参考品牌仍只有三只松鼠、百草味、良品铺子。

## 主图比例

详情屏线框对应工单里的 9:16，设计尺寸 900×1600 px。五张主图线框也按 900×1600 截出，只为和 Excel 参考图列对齐。真实主图比例待按销售渠道确认，这里没有确认任何渠道尺寸。工单 B4 和主图 J 列说明写了同一句。

## 目录里有什么

| 文件 | 作用 |
|---|---|
| `detail-request-with-wireframe.xlsx` | 正式三图工单加 J 列「本屏线框」。36 张参考图、36 条商品链接，另有 12 张线框 |
| `preview-with-wireframe.jpg` | 前 4 屏预览：文案、三张参考图、右侧线框 |
| `angus-wireframe.html` | 可滑动的黑白线框长图。样式写在文件内，不引用外部图片、脚本或字体文件 |
| `angus-wireframe-fullpage.png` | 上述长图的整页截图 |
| `screens/` | 12 张 900×1600 线框，供重新嵌入 J 列 |
| `build_wireframe.py` `add_j_column.py` `make_preview.py` `shoot.sh` | 可复现脚本 |
| `_shot_full.html` `_shots/s01.html`–`s12.html` `manifest.json` | 截图用的小页面和屏序 |

不纳入本目录：`_shots/*.raw.png`（中间截图）、`source-copy.xlsx`（与正式工单重复）、字体文件。

## 重新生成

在本目录执行。未传路径时，优先使用旁边的 `source-copy.xlsx`；没有这份副本时，读取 `../angus-beef/detail-request.xlsx`。脚本不会改写正式工单。

```sh
python3 build_wireframe.py
bash shoot.sh
python3 add_j_column.py
python3 make_preview.py
```

`shoot.sh` 需要本机的 headless Chrome，以及已安装的思源黑体 / Noto Sans CJK，用来按已审核的线框重新截图。`make_preview.py` 重做预览图时也会查找本机中文字体，优先 Noto Sans CJK。两处都不附带字体文件。中间文件 `_shots/*.raw.png` 可以删掉。

已提交的 HTML、Excel、预览图和 `screens/` 就是这次试看的结果。没有中文字体或 Chrome 时，直接使用这些文件即可。
