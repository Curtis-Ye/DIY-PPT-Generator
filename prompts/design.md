# Slide Designer

## Layout Types

### TITLE — 封面页
- 标题居中，使用 `cover_title` 字体，`cover_title_size` 字号
- 副标题居中，使用 `cover_subtitle` 字体
- 作者/日期底部居中，使用 `cover_info` 字体
- 背景使用主题色 `primary`

### TOC — 目录页
- 标题"目录"居左，使用 `heading` 字体
- 目录项竖排，使用 `body` 字体，编号+标题
- 标题下方加一条主题色分隔线

### TEXT — 纯文字页
- 标题居左，使用 `heading` 字体，粗体
- 内容区：bullet 列表，使用 `body` 字体
- 标题下方加一条主题色分隔线

### IMAGE_TEXT — 左图右文
- 左侧 50% 宽度：图片占位区
- 右侧 50% 宽度：标题 + bullet 列表
- 图片区域显示提示文字"【配图：{image_suggestion}】"，使用 `caption` 字体

### TEXT_IMAGE — 左文右图
- 与 IMAGE_TEXT 镜像布局
- 左侧文字，右侧图片

### RESULT — 数据/结果页
- 标题居左，使用 `heading` 字体
- 图表占位区占页面上方 65%
- 下方 35%：简短结论 bullet，使用 `body` 字体

### FLOWCHART — 流程图/架构页
- 标题居左，使用 `heading` 字体
- 流程图/架构图占位区居中，占 75% 面积
- 底部可有简短的步骤说明

### SECTION — 章节分隔页
- 章节编号大号显示（`section_number_size`）
- 章节标题大字居中，使用 `heading` 字体，粗体
- 纯色背景或大面积色块（`primary`）
- 可选副标题说明本节内容

### END — 结束页
- "谢谢"居中，使用 `end_title_size` 字号，粗体
- 副标题居中
- 简洁大方

## 配色规则

- 标题统一使用主题主色 `primary`
- 正文使用 `text_dark`
- 分隔线/强调元素使用 `primary` 或 `accent`
- 背景使用 `background`，不可过暗
- 图片占位区使用 `secondary` 填充

## 字体层级

用户可在 `config.yaml` 中为每个层级独立选择字体和字号：

| 配置项 | 用途 | 默认字体 | 默认字号 |
|--------|------|----------|----------|
| `cover_title` | 封面大标题 | 微软雅黑 | 40pt |
| `cover_subtitle` | 封面副标题 | 微软雅黑 | 20pt |
| `cover_info` | 封面作者/日期 | 微软雅黑 | 14pt |
| `heading` | 内容页标题 | 微软雅黑 | 32pt |
| `body` | 正文 bullet | 微软雅黑 | 18pt |
| `caption` | 图表标注 / 占位提示 | 微软雅黑 | 12pt |
| `page_number` | 页码 | 微软雅黑 | 10pt |
| `section_number_size` | 章节分隔页数字 | — | 72pt |
| `end_title_size` | 结束页文字 | — | 44pt |

### 常用字体推荐

Windows:
- 微软雅黑 — 现代简洁，通用性强
- 黑体 — 庄重醒目，适合标题
- 宋体 — 传统正式，适合正文
- 楷体 — 学术论文常用
- Arial / Calibri — 英文内容
