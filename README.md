# 左联人际关系网络数据可视化

> **大创项目** — 基于多源文献的中国左翼作家联盟（左联）人际关系网络数据提取与可视化研究

## 📖 项目简介

本项目旨在从多种历史文献中系统性提取中国左翼作家联盟（1930-1936）成员间的人际关系数据，构建社会网络分析数据集。项目涵盖 **150+ 位左联相关历史人物**，通过 OCR 识别、自然语言处理和人物共现分析等方法，从一手史料中挖掘人物间的交往关系。

## 📂 项目结构

```
├── 数据提取脚本
│   ├── extract_relationships.py    # 从《鲁迅日记》中提取人际关系
│   ├── extract_zuolian.py          # 从《左联回忆录》PDF中提取人际关系
│   ├── extract_from_cidian_shi.py  # 从《左联词典》和《左联史》TXT中提取关系
│   └── aggregate_luxun_data.py     # 鲁迅节点数据清洗与按月聚合
│
├── 数据处理工具
│   ├── convert_to_txt.py           # OCR识别PDF并转换为TXT文件
│   ├── ocr_search.py               # OCR识别《左联词典》并检索成员
│   ├── find_isolated_members.py    # 识别孤岛成员（无关系连接的节点）
│   ├── search_isolated_v2.py       # 孤岛成员检索（改进版）
│   ├── filter_luxun_diary.py       # 筛选《鲁迅日记》中与左联相关的内容
│   ├── export_isolated.py          # 导出孤岛成员数据
│   ├── fix_data.py                 # 数据修复工具
│   ├── fix_birth_death.py          # 修复生卒年数据
│   └── transfer_data.py            # 数据迁移工具
│
├── 数据文件
│   ├── 大创数据收集1.xlsx           # 主数据文件（Sheet1: 人物信息, Sheet2: 关系数据）
│   ├── 孤岛成员名单.xlsx            # 孤岛成员清单
│   ├── isolated_members_found.xlsx  # 已找到的孤岛成员
│   └── isolated_members_ocr_result.xlsx  # OCR检索结果
│
└── 文本资料
    ├── 日记全编：全2册 (鲁迅 著).txt  # 《鲁迅日记》全文
    ├── 左联词典.txt                   # 《左联词典》OCR文本
    └── 左联史.txt                     # 《左联史》OCR文本
```

## 📊 数据结构

### Sheet1 — 人物信息表

| 字段 | 说明 |
|------|------|
| `Entity_ID` | 唯一标识符（如 `ZLH-001`） |
| `True_Name` | 真实姓名 |
| `Alias` | 别名/笔名 |
| `Birth_Year` / `Death_Year` | 生卒年 |

### Sheet2 — 关系数据表

| 字段 | 说明 |
|------|------|
| `Source_ID` | 关系发起方 ID |
| `Target_ID` | 关系接收方 ID |
| `Relation_Type` | 关系类型（强关联-亲属/组织、弱关联-通信/时空共现） |
| `Context` | 关系描述上下文 |
| `Evidence_Ref` | 文献出处 |
| `Weight` | 关系权重（1-10） |

## 📚 数据来源

| 文献 | 类型 | 提取方式 |
|------|------|----------|
| 《鲁迅日记》 | 文本 | 正则表达式 + 人物别名匹配 |
| 《左联回忆录》 | 扫描PDF | pdfplumber 文本提取 + NLP |
| 《左联词典》（姚辛） | 扫描PDF | Tesseract OCR + 人物共现分析 |
| 《左联史》（姚辛） | 扫描PDF | Tesseract OCR + 人物共现分析 |

## 🛠️ 技术栈

- **Python 3.x**
- **OCR**: Tesseract-OCR + pytesseract + pdf2image
- **PDF解析**: pdfplumber
- **数据处理**: pandas, openpyxl
- **文本分析**: re（正则表达式）, 人物别名映射表

## 🚀 快速开始

### 环境配置

```bash
# 安装 Python 依赖
pip install pandas openpyxl pdfplumber pytesseract pdf2image

# 安装 Tesseract-OCR（Windows）
# 下载: https://github.com/tesseract-ocr/tesseract
# 安装后确保 tesseract.exe 在 PATH 中
```

### 数据提取流程

```mermaid
graph LR
    A[原始PDF/TXT] --> B[OCR识别]
    B --> C[文本预处理]
    C --> D[人物识别]
    D --> E[关系提取]
    E --> F[数据清洗]
    F --> G[Excel输出]
```

1. **OCR 转换**（如需处理新 PDF）

   ```bash
   python convert_to_txt.py
   ```

2. **提取关系数据**

   ```bash
   python extract_relationships.py        # 鲁迅日记
   python extract_zuolian.py              # 左联回忆录
   python extract_from_cidian_shi.py      # 左联词典 & 左联史
   ```

3. **数据清洗与聚合**

   ```bash
   python aggregate_luxun_data.py         # 聚合鲁迅通信记录
   python filter_luxun_diary.py           # 筛选核心关系
   ```

4. **孤岛成员处理**

   ```bash
   python find_isolated_members.py        # 识别无连接的成员
   python ocr_search.py                   # 在文献中检索孤岛成员
   ```

## 📝 注意事项

- 大型 PDF 文件（左联词典、左联史等）和 OCR 中间文件未包含在仓库中（受 GitHub 文件大小限制），请自行准备原始 PDF
- OCR 结果可能存在识别误差，建议人工校验关键数据
- `chi_sim.traineddata`（中文训练数据）需放在项目根目录下供 Tesseract 使用

## 📄 License

本项目仅用于学术研究，文献资料版权归原作者所有。
