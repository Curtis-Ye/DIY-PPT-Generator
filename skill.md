# DIY PPT Generator

## Description

根据用户提供的文档（.docx / .pdf），智能生成演示文稿（PPT）。适用于技术报告、毕业答辩、科研汇报、商业提案等场景。

## Workflow

按以下步骤依次执行，每步完成后向用户汇报进展。

### Step 0: 读取用户需求

检查 `examples/request.yaml` 是否存在。如果存在，读取其中的用户偏好：

- `page_count`: 目标页数（用户指定的优先级最高，覆盖 prompt 中的默认值）
- `document`: 要处理的文档路径
- `meta`: 演示文稿的标题/副标题/作者/日期
- `theme`: 主题选择
- `notes`: 特殊要求

如果该文件不存在，使用默认值（15 页），并向用户确认文档路径。

### Step 1: 读取文档

调用 read_doc 工具提取文档内容和结构：

```bash
python tools/read_doc.py <document_path>
```

工具会输出 JSON，包含：
- `title`: 文档标题
- `headings`: 章节层级结构
- `paragraphs`: 正文段落列表
- `tables`: 表格数据（如有）

### Step 2: 生成大纲

根据文档结构，参考 `prompts/outline.md` 的规则生成 PPT 大纲。

**页数**: 优先使用 `request.yaml` 中用户指定的 `page_count`。如果用户未指定，默认 15 页。

输出格式：
```json
[
  { "page": 1, "title": "封面" },
  { "page": 2, "title": "目录" },
  ...
]
```

约束：
- 总页数以用户指定的 `page_count` 为准
- 必须包含：封面、目录、核心内容（40%-50% 页数）、总结

### Step 3: 为每页生成内容

参考 `prompts/summary.md` 的规则，为大纲中的每一页提取/压缩内容。

对每页输出：
- `bullets`: 要点列表（3-6条，每条不超过20字）
- `speaker_notes`: 讲稿（供演讲者参考，1-3句话）
- `image_suggestion`: 配图建议（关键词描述，用于搜索或生成配图）

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

### Step 5: 生成 PPT 文件

将完整数据组装为 JSON 并写入临时文件，然后调用：

```bash
python tools/create_ppt.py <slide_data.json> -c config.yaml -o <output.pptx>
```

JSON 结构见下方 Output Format。

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
      "image_suggestion": "配图关键词描述",
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
- 封面和结束页必须简洁

---

## Config

`config.yaml` 控制幻灯片尺寸、主题色、字体等参数。生成 PPT 时读取该配置。用户可通过修改该文件自定义样式。
