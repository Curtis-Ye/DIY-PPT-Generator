# DIY PPT Generator — 使用指南

根据 Word 文档（.docx）或 PDF 一键生成演示文稿。

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 放入文档

把你的文档放到 `examples/` 目录下：

```
examples/
  ├── 我的报告.docx    ← 你的文档放这里
  ├── request.yaml     ← 你的需求写这里
  └── ...
```

### 3. 填写需求

编辑 `examples/request.yaml`：

```yaml
# 要生成几页？（8~30，建议 12~18）
page_count: 15

# 文档路径
document: examples/我的报告.docx

# 演示文稿信息（选填，留空自动提取）
meta:
  title: ""
  subtitle: "硕士学位论文答辩"
  author: "张三"
  date: ""

# 主题：blue / dark / green
theme: blue

# 额外要求（选填）
notes: ""
```

### 4. 对 AI 说一句话

在 Claude Code 中输入：

> "帮我用 examples/我的报告.docx 生成一份 PPT"

或者：

> "读取 examples/request.yaml，生成 PPT"

---

## AI 工作流程

AI 会按以下步骤自动执行，你只需等待结果：

| 步骤   | 说明                                                    |
| ------ | ------------------------------------------------------- |
| Step 0 | 读取 `request.yaml` 中的页数、主题、特殊要求            |
| Step 1 | 调用 `read_doc.py` 提取文档标题、章节、正文             |
| Step 2 | 根据文档结构和用户指定页数生成大纲                      |
| Step 3 | 将每章内容压缩为 3-6 条要点 + 讲稿 + 配图建议           |
| Step 4 | 为每页确定版式（纯文字 / 图文混排 / 数据图表 / 流程图） |
| Step 5 | 调用 `create_ppt.py` 生成 `.pptx` 文件                  |
| Step 6 | 验证页数和内容，报告结果                                |

---

## 自定义

### 换主题

编辑 `examples/request.yaml`：

```yaml
theme: dark
```

可选：`blue`（默认）/ `dark` / `green`。也可在 `config.yaml` 中自定义任意配色。

### 指定页数

```yaml
page_count: 10    # 想要 10 页就写 10
```

### 特殊要求

```yaml
notes: "重点突出实验数据，弱化公式推导，多放架构图"
```

AI 会在生成大纲和内容时遵循这些要求。

---

## 目录结构

```
PPT_skill/
├── skill.md              # Skill 定义，AI 的工作流程
├── config.yaml           # 幻灯片尺寸、配色、字体配置
├── requirements.txt      # Python 依赖
├── README.md             # 本文件
├── prompts/              # 各阶段的 Prompt 规则
│   ├── outline.md        #   大纲生成规则
│   ├── summary.md        #   内容压缩规则
│   └── design.md         #   版式设计规则
├── tools/                # 可执行工具脚本
│   ├── read_doc.py       #   文档读取（.docx / .pdf）
│   └── create_ppt.py     #   PPT 文件生成
└── examples/             # 用户文件
    ├── request.yaml      #   ← 你的需求写这里
    ├── input.docx        #   ← 你的文档放这里
    ├── output.json        #   中间产物（幻灯片数据）
    └── output.pptx        #   最终产物（生成的 PPT）
```

---

## 生成的 PPT 结构

```json
{
  "meta": { "title": "...", "author": "...", "date": "..." },
  "slides": [
    { "page": 1,  "layout": "TITLE",      "title": "封面" },
    { "page": 2,  "layout": "TOC",        "bullets": ["目录项", ...] },
    { "page": 3,  "layout": "IMAGE_TEXT", "bullets": [...], "image_suggestion": "配图关键词" },
    { "page": 15, "layout": "END",        "title": "谢谢" }
  ]
}
```

9 种版式：`TITLE` / `TOC` / `TEXT` / `IMAGE_TEXT` / `TEXT_IMAGE` / `RESULT` / `FLOWCHART` / `SECTION` / `END`。

---

## 常见问题

**Q: 支持 .doc 格式吗？**
A: 不支持旧版 .doc。请用 Word 另存为 .docx。

**Q: 配图会自动生成吗？**
A: 不会。AI 会给出配图建议关键词（写在占位区），你需要手动替换为实际图片。

**Q: 生成效果不满意怎么办？**
A: 编辑 `examples/request.yaml` 中的 `notes` 字段写明要求，重新生成。也可以直接修改 `prompts/` 目录下的规则文件。
