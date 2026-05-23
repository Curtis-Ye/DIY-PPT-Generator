# PPT 模板使用说明

将你的 `.pptx` 模板文件放入此目录。

## 模板是什么

模板是一个预先设计好版式、配色、字体的 `.pptx` 文件。使用模板后，生成的 PPT 会继承模板的视觉风格（背景、配色、字体、slide layout）。

## 如何制作模板

1. 在 PowerPoint 中新建文件
2. 进入「视图 → 幻灯片母版」
3. 按需设计以下版式（至少包含一个空白版式）：
   - 封面版式 — 对应 TITLE
   - 正文版式 — 对应 TEXT / TOC
   - 图文版式 — 对应 IMAGE_TEXT / TEXT_IMAGE
   - 数据版式 — 对应 RESULT
   - 流程图版式 — 对应 FLOWCHART
   - 分隔版式 — 对应 SECTION
   - 结束版式 — 对应 END
4. 关闭母版视图，保存为 `.pptx`

## 如何使用模板

在 `examples/request.yaml` 中指定模板路径：

```yaml
template: templates/my_theme.pptx
```

如果模板中的 slide layout 编号与默认不同，在 `config.yaml` 中配置映射：

```yaml
template:
  path: templates/my_theme.pptx
  layout_mapping:
    TITLE: 0
    TEXT: 1
    IMAGE_TEXT: 2
    RESULT: 3
    END: 8
```

未列出的版式使用默认映射。

## 没有模板时

如果不指定模板（`template: ""`），AI 会使用内置的程序化样式，通过 `config.yaml` 控制配色和字体。
