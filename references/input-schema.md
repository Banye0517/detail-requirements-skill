# 导出器数据格式

智能体调研、查看图片并写好文案后，生成UTF-8 JSON。图片先保存到本次工作目录。路径可用绝对路径，也可相对于JSON文件的相对路径。图片已经裁成9:16，建议720×1280或900×1600。下面是字段示意，正式输入需替换示意内容并为每屏填足三条references。

```json
{
  "product_name": "本品名称",
  "selling_points": ["用户原始卖点一", "用户原始卖点二"],
  "audience": "儿童零食，具体适龄待确认",
  "summary": "运营审核草稿。原料、营养表与包装待核对。",
  "screens": [
    {
      "id": "第1屏 / 原表U10",
      "task": "产品识别",
      "headline": "小太阳里藏着米香",
      "subtitle": "太阳花造型轻酥米饼",
      "labels": "黑土好米｜太阳花造型（用户主张待核对）",
      "tips": "产品名称另列；品牌标志待提供。",
      "visual": "单片太阳花米饼作视觉中心，包装在下方。上方留标题，侧光表现纹理。理由为先识别形态，再读卖点。",
      "pending": "本品原图、品牌与包装待补。",
      "references": [
        {
          "image": "images/reference-01.jpg",
          "brand": "指定名单中的实际品牌",
          "product_name": "实际商品名称",
          "product_url": "https://example.com/actual-product",
          "borrow": "产品与标题的主次关系",
          "application": "本品换成太阳花实物，标题位置留白",
          "image_url": "https://example.com/actual-detail-image.jpg",
          "crop_note": "来自详情正文；原图裁切坐标x,y,w,h；如有补边写明"
        }
      ]
    }
  ],
  "main_images": [],
  "research": [
    {
      "brand": "实际品牌",
      "product_name": "实际商品",
      "product_url": "https://example.com/actual-product",
      "observed": "实际看到的文案、图片和屏序",
      "application": "本品借鉴的具体方法",
      "limits": "店铺身份、可见范围、销量证据或其缺口"
    }
  ],
  "brand_search": [
    {
      "brand": "本次应覆盖的品牌",
      "status": "已检索，实际采用或未采用及原因",
      "source": "https://example.com/actual-source"
    }
  ],
  "pending": [
    {
      "owner": "运营 / 产品 / 设计",
      "item": "钙含量",
      "required": "本品营养成分表及宣称依据",
      "reason": "决定高钙屏能否定稿"
    }
  ]
}
```

`main_images`中的对象与`screens`同结构。每个对象的references必须为三条，每条包含实际存在的图片文件、品牌、商品名和商品URL。示意中的example.com不能用于交付。校验脚本只检查URL结构，智能体负责访问并核对真实性与对应关系。

`image_url`与`crop_note`用于保留来源和裁切信息，图片本体由`image`提供。品牌、商品名及商品链接在Excel图片正下方显示，借鉴和应用在再下一行。每条图片引用都计入验收，包括主图建议行。

运行方式

```sh
python scripts/build_workbook.py input.json output.xlsx
```

缺图草案允许显式使用`--allow-incomplete`，Excel会标出缺口并报告未完整。正式交付验收仍要求每屏三张。已有文件默认拒绝覆盖，需要时明确指定`--overwrite`。

重新生成空白模板

```sh
python scripts/build_workbook.py blank-template assets/detail-request-template.xlsx
```

空白模板包含7屏和5张主图的默认字段。实际卖点数量不同可调整屏数。模板中空白是供填报的图位，不能作为已有参考图。
