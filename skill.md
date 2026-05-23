# DIY PPT Generator

## Description

根据用户提供的文档（.docx / .pdf），智能生成演示文稿（PPT）。支持自定义 PPT 模板和逐级字体配置。

## Workflow

按以下步骤依次执行，每步完成后向用户汇报进展。

### Step 0: 读取用户需求

检查 `examples/request.yaml` 是否存在。如果存在，读取其中的用户偏好：

- `page_count`: 目标页数（优先级最高）
- `document`: 要处理的文档路径
- `template`: PPT 模板路径（留空则使用内置样式）
- `meta`: 标题/副标题/作者/日期
- `theme`: 主题选择（使用模板时此项被模板覆盖）
- `font`: 字体自定义（选填，覆盖 config.yaml 默认值）
- `notes`: 特殊要求

如果该文件不存在，使用默认值（15 页），并向用户确认文档路径。

### Step 1: 读取文档

调用 read_doc 工具：

```bash
python tools/read_doc.py <document_path>
```

输出 JSON：`title`, `headings`, `paragraphs`, `tables`。

### Step 2: 生成大纲

根据文档结构，参考 `prompts/outline.md` 的规则生成 PPT 大纲。

**页数**: 以 `request.yaml` 中的 `page_count` 为准，默认 15 页。

输出格式：
```json
[
  { "page": 1, "title": "封面", "source_sections": ["文档标题"] },
  { "page": 2, "title": "目录", "source_sections": [] }
]
```

约束：
- 总页数以用户指定的 `page_count` 为准
- 必须包含：封面、目录、核心内容（40%-50% 页数）、总结

### Step 3: 为每页生成内容

参考 `prompts/summary.md` 的规则，为大纲中的每一页提取/压缩内容。

对每页输出：
- `bullets`: 要点列表（3-6条，每条不超过20字）
- `speaker_notes`: 讲稿（1-3句话）
- `image_suggestion`: 配图建议关键词

### Step 4: 确定版式

参考 `prompts/design.md` 的规则，为每页指定 layout 类型：

| layout | 适用场景 |
|--------|---------|
| TITLE | 封面页 |
| TOC | 目录页 |
| TEXT | 纯文字要点 |
| IMAGE_TEXT | 左图右文 |
| TEXT_IMAGE | 左文右图 |
| RESULT | 图表为主+简短结论 |
| FLOWCHART | 流程图/架构图 |
| SECTION | 章节分隔页 |
| END | 结束页 |

如果用户指定了模板，版式会自动匹配模板中的 slide layout。

### Step 5: 生成 PPT 文件

将完整数据组装为 JSON 写入临时文件，然后调用：

```bash
python tools/create_ppt.py <slide_data.json> -c config.yaml -o <output.pptx>
```

工具会自动：
- 如果 `request.yaml` 中指定了模板路径，使用该模板的版式和主题
- 如果没有模板，使用内置的程序化样式
- 应用 `config.yaml` 中的字体配置（封面标题、内容标题、正文等分别可设）

### Step 6: 验证与交付

- 检查 PPT 页数、内容完整性
- 告知用户生成结果及文件路径

---

## Output Format

```json
{
  "meta": {
    "title": "演示文稿标题",
    "author": "作者",
    "date": "2026-05-23"
  },
  "slides": [
    {
      "page": 1,
      "title": "封面标题",
      "layout": "TITLE",
      "subtitle": "副标题",
      "author": "作者",
      "date": "日期"
    },
    {
      "page": 2,
      "title": "目录",
      "layout": "TOC",
      "bullets": ["背景与意义", "方案设计", "实验验证", "总结展望"]
    },
    {
      "page": 3,
      "title": "研究背景",
      "layout": "IMAGE_TEXT",
      "bullets": ["要点一", "要点二", "要点三"],
      "image_suggestion": "配图关键词",
      "speaker_notes": "本页讲稿内容..."
    },
    {
      "page": 15,
      "title": "谢谢",
      "layout": "END",
      "subtitle": "请各位老师批评指正"
    }
  ]
}
```

---

## Rules

- 每页最多 6 个要点
- 每个要点不超过 20 字
- 禁止出现大段文字段落
- 优先图文混排（IMAGE_TEXT / TEXT_IMAGE / RESULT）
- 实验/数据结果必须使用 RESULT 版式
- 总结页不超过 3 条结论

---

## Config

| 文件 | 用途 |
|------|------|
| `config.yaml` | 幻灯片尺寸、主题色、字体层级、模板版式映射 |
| `examples/request.yaml` | 页数、文档路径、模板选择、特殊要求 |
| `templates/` | 用户自定义 .pptx 模板存放目录 |
