# 文档生成技能使用指南

## 📋 目录
- [快速开始](#快速开始)
- [多级标题](#多级标题)
- [表格功能](#表格功能)
- [图片功能](#图片功能)
- [格式设置](#格式设置)
- [完整示例](#完整示例)
- [命令行参数](#命令行参数)

---

## 快速开始

### 1. 基本用法

```bash
# 创建简单的Word文档
python3 create_word.py "my_doc.docx" \
  --title "文档标题" \
  --content "这是文档内容"
```

### 2. 从JSON文件创建富文档

```bash
python3 create_word.py "rich_document.docx" \
  --title "我的报告" \
  --content-file "content.json"
```

---

## 多级标题

### 标题级别

Word文档支持多级标题结构：

| 级别 | 示例 | 用途 |
|------|------|------|
| 0 | 主标题 | 文档主标题 |
| 1 | 一级标题 | 主要章节 |
| 2 | 二级标题 | 子章节 |
| 3-9 | 更细粒度 | 小节、细项 |

### JSON示例

```json
{
  "sections": [
    {
      "heading": "第一章：概述",
      "level": 1,
      "content": [...]
    },
    {
      "heading": "1.1 背景介绍",
      "level": 2,
      "content": [...]
    },
    {
      "heading": "1.1.1 历史沿革",
      "level": 3,
      "content": [...]
    }
  ]
}
```

---

## 表格功能

### 创建表格

表格支持表头、数据行和自定义样式：

```json
{
  "type": "table",
  "headers": ["姓名", "年龄", "职业"],
  "data": [
    ["张三", "25", "工程师"],
    ["李四", "30", "设计师"],
    ["王五", "28", "产品经理"]
  ],
  "style": "Light Grid Accent 1",
  "autofit": true
}
```

### 表格样式

可选样式（基于Word内置样式）：

- `Light Grid Accent 1` - 浅色网格
- `Light Shading Accent 2` - 浅色阴影
- `Medium Grid 1` - 中等网格
- `Table Grid` - 标准网格
- `No Style` - 无样式

### 表格属性

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| headers | array | 否 | 表头列表 |
| data | array of arrays | 否 | 表格数据行 |
| style | string | 否 | 表格样式，默认 'Light Grid Accent 1' |
| autofit | boolean | 否 | 是否自动调整列宽，默认 true |

---

## 图片功能

### 插入图片

```json
{
  "type": "image",
  "path": "/path/to/image.png",
  "width": 5,
  "caption": "图片说明文字",
  "alignment": "center"
}
```

### 图片属性

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| path | string | 是 | 图片文件路径（支持PNG、JPG等） |
| width | number | 否 | 图片宽度（英寸），默认 5 |
| caption | string | 否 | 图片说明（标题） |
| alignment | string | 否 | 对齐方式：'left'、'center'、'right' |

### 示例

```json
{
  "type": "image",
  "path": "architecture_diagram.png",
  "width": 6,
  "caption": "系统架构设计图",
  "alignment": "center"
}
```

---

## 格式设置

### 段落格式

```json
{
  "type": "paragraph",
  "text": "这是段落文本",
  "bold": true,
  "italic": false,
  "underline": true,
  "font_size": 16,
  "font_color": "FF0000",
  "alignment": "center"
}
```

### 格式属性

| 字段 | 类型 | 说明 |
|------|------|------|
| bold | boolean | 粗体 |
| italic | boolean | 斜体 |
| underline | boolean | 下划线 |
| font_size | number | 字体大小（磅） |
| font_color | string | 字体颜色（十六进制，如 "FF0000"） |
| alignment | string | 对齐：'left'、'center'、'right'、'justify' |

### 对齐方式

| 值 | 效果 |
|-----|------|
| left | 左对齐 |
| center | 居中 |
| right | 右对齐 |
| justify | 两端对齐 |

---

## 完整示例

### 完整JSON结构示例

```json
{
  "sections": [
    {
      "heading": "项目报告",
      "level": 1,
      "content": [
        {
          "type": "paragraph",
          "text": "项目提交日期：2026年3月",
          "italic": true,
          "alignment": "right"
        },
        {
          "type": "paragraph",
          "text": "这是一个功能完整的文档示例，展示了所有支持的格式。"
        },
        {
          "type": "heading",
          "text": "1. 项目背景",
          "level": 2
        },
        {
          "type": "list",
          "list_type": "bullet",
          "items": [
            "目标是创建强大的文档生成系统",
            "支持多种格式和丰富功能",
            "易于集成和使用"
          ]
        },
        {
          "type": "table",
          "headers": ["功能模块", "负责人", "状态"],
          "data": [
            ["文档解析", "高松灯", "✅ 完成"],
            ["模板系统", "开发中", "🔄 50%"],
            ["测试套件", "待启动", "⏳ 计划中"]
          ]
        },
        {
          "type": "image",
          "path": "/path/to/project_photo.jpg",
          "width": 5,
          "caption": "项目团队合影",
          "alignment": "center"
        }
      ]
    }
  ]
}
```

---

## 命令行参数

### create_word.py

```
用法: python3 create_word.py <输出文件> [选项]

必需参数:
  <输出文件>         输出.docx文件路径

可选参数:
  --title TITLE       文档标题
  --subtitle TEXT     文档副标题
  --content TEXT      JSON字符串或文本内容
  --content-file PATH 包含内容的JSON/文本文件路径
  --help             显示此帮助信息

示例:
  python3 create_word.py "report.docx" --title "我的报告" --content-file "data.json"
  python3 create_word.py "note.docx" --title "笔记" --content "简单文本内容"
```

---

## 常见问题

### Q: 图片路径不支持？
A: 确保路径正确且文件存在。支持相对路径和绝对路径。

### Q: 表格数据显示异常？
A: 检查数据格式，应为数组的数组（`[[row1], [row2], ...]`），每行长度建议一致。

### Q: 如何创建超过3级的标题？
A: 在JSON中设置 `"level": 4`（或其他数字1-9），Word支持最多9级标题。

### Q: 字体颜色如何指定？
A: 使用十六进制颜色代码，如 `"FF0000"`（红色）、`"00FF00"`（绿色）、`"0000FF"`（蓝色）。

---

## 版本历史

- **v2.0** (2026-03-15): 新增多级标题、表格、图片、富文本格式支持
- **v1.0**: 基础文档创建功能

---

**需要帮助？** 查看 `demo_example.py` 运行示例，或联系高松灯 🐧
