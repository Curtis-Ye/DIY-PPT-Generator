# DIY PPT Generator — 使用指南

根据 Word 文档（.docx）或 PDF 一键生成演示文稿。支持自定义模板和逐级字体配置。

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 放入文档

把你的文档放到 `examples/` 目录下。

### 3. 填写需求

编辑 `examples/request.yaml`：

```yaml
# 要生成几页？
page_count: 15

# 文档路径
document: examples/我的报告.docx

# PPT 模板（选填，留空 = 内置样式）
template: ""

# 演示文稿信息（选填）
meta:
  subtitle: "硕士学位论文答辩"
  author: "张三"

# 主题：blue / dark / green
theme: blue

# 额外要求（选填）
notes: ""
```

### 4. 对 AI 说一句话

> "帮我用 examples/我的报告.docx 生成一份 PPT"

---

## 使用自定义模板

将你公司的 `.pptx` 模板放入 `templates/` 目录，然后在 `request.yaml` 中指定：

```yaml
template: templates/公司模板.pptx
```

生成的 PPT 会自动继承模板的版式、配色和字体风格。

详见 [templates/README.md](templates/README.md)。

---

## 自定义字体

`config.yaml` 中每个层级的字体和字号都可以独立设置：

```yaml
font:
  # 字体名称
  cover_title: 微软雅黑       # 封面大标题
  cover_subtitle: 微软雅黑    # 封面副标题
  cover_info: 微软雅黑        # 封面作者/日期
  heading: 微软雅黑           # 内容页标题
  body: 微软雅黑              # 正文 bullet
  caption: 微软雅黑           # 图表标注 / 占位提示
  page_number: 微软雅黑       # 页码

  # 字号 (pt)
  cover_title_size: 40
  heading_size: 32
  body_size: 18
  # ...
```

也可以临时覆盖（在 `request.yaml` 中）：

```yaml
font:
  heading: 黑体
  body: 宋体
```

---

## AI 工作流程

| 步骤   | 说明                                                    |
| ------ | ------------------------------------------------------- |
| Step 0 | 读取 `request.yaml` 中的页数、模板、主题、特殊要求      |
| Step 1 | 调用 `read_doc.py` 提取文档标题、章节、正文             |
| Step 2 | 根据文档结构和用户指定页数生成大纲                      |
| Step 3 | 将每章内容压缩为 3-6 条要点 + 讲稿 + 配图建议           |
| Step 4 | 为每页确定版式（纯文字 / 图文混排 / 数据图表 / 流程图） |
| Step 5 | 调用 `create_ppt.py` 生成 `.pptx`（应用模板和字体配置） |
| Step 6 | 验证页数和内容，报告结果                                |

---

## 目录结构

```
PPT_skill/
├── skill.md              # Skill 定义，AI 的工作流程
├── config.yaml           # 尺寸、配色、字体层级、模板版式映射
├── requirements.txt      # Python 依赖
├── README.md             # 本文件
├── prompts/              # 各阶段 Prompt 规则
│   ├── outline.md        #   大纲生成规则
│   ├── summary.md        #   内容压缩规则
│   └── design.md         #   版式设计与字体层级
├── tools/                # 可执行工具脚本
│   ├── read_doc.py       #   文档读取（.docx / .pdf）
│   └── create_ppt.py     #   PPT 文件生成（支持模板）
├── templates/            # 用户 PPT 模板
│   └── README.md         #   模板制作与使用说明
└── examples/             # 用户文件
    ├── request.yaml      #   ← 你的需求写这里
    ├── input.docx        #   ← 你的文档放这里
    ├── output.json        #   中间产物（幻灯片数据）
    └── output.pptx        #   最终产物（生成的 PPT）
```

---

## 常见问题

**Q: 支持 .doc 格式吗？**
A: 不支持旧版 .doc，请用 Word 另存为 .docx。

**Q: 如何制作模板？**
A: 在 PowerPoint 中进入「视图 → 幻灯片母版」设计版式，保存为 `.pptx` 放入 `templates/` 目录。详见 [templates/README.md](templates/README.md)。

**Q: 使用模板后字体配置还有效吗？**
A: AI 仍会在模板基础上叠加 config.yaml 的字体设置。如果模板已定义好字体，config 中的设置会自动生效。

**Q: 配图会自动生成吗？**
A: AI 在占位区给出配图建议关键词，你需要手动替换为实际图片。

**Q: 效果不满意怎么办？**
A: 编辑 `request.yaml` 的 `notes` 字段或调整 `prompts/` 目录下的规则文件后重新生成。
