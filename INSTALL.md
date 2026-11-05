# 安装与使用

界面名称为「详情需求表」，通用技能标识为 `detail-requirements`。文件夹保留英文名，便于符合常见Skill加载规则。

## 支持技能文件的智能体

将ZIP完整解压，把 `detail-requirements` 整个文件夹导入平台技能管理器，或放入该平台文档规定的技能目录。入口为 `SKILL.md`。部分平台支持上传ZIP，按其导入提示操作。平台目录和调用语法各不相同，以平台说明为准。

## 只支持自定义指令或知识文件的智能体

把 `SKILL.md` 作为工作流指令，并上传 `references/`、`assets/` 和 `scripts/`。设置显示名称为「详情需求表」。若无法加载目录，上传这些文件并指示智能体按文件名读取引用。

## 运行要求

完整交付需要联网检索、查看和获取真实详情图、生成带图片及超链接的Excel，以及下载最终文件的能力。Skill指令不能为宿主增加其未提供的工具。已制作通用文件包，尚未逐个平台实机安装验证。

支持Python的环境可运行附带导出器。推荐使用宿主已提供的表格工具与依赖。需要自行准备依赖时使用官方Python包仓库。

```sh
python -m pip install openpyxl Pillow
python scripts/build_workbook.py input.json output.xlsx
```

有uv的环境也可临时加载依赖。

```sh
uv run --with openpyxl --with pillow python scripts/build_workbook.py input.json output.xlsx
```

JSON由智能体按 `references/input-schema.md` 在调研后生成。导出器不会凭产品名自行调用AI，也不会自行搜索图片。只输入名称和卖点即可完成，指的是具备上述能力的智能体执行完整Skill流程。

## 日常输入示例

```text
用详情需求表。
产品名称：轻酥太阳米饼，儿童零食。
卖点：黑土好米，米香清甜；芝香蛋浓，果蔬真味；高钙配方；太阳花造型。
再加一个使用场景。
```

自然语言也可以。仅发名称和卖点即可启动，年龄段等关键信息未知时，智能体会标注假设或集中提问。其他缺失资料集中放进待补清单。
