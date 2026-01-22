# [中文](README.md) | [English](README_EN.md)

# 复杂 Excel 解析器 (Complex Excel Parser)

这是一个生产级的 Python 应用程序，用于解析复杂结构甚至**受损**的 Excel 文件。它集成了全域基因修复技术、Qwen LLM 语义提取以及多种结构化解析模式。

## 核心功能

### 1. 全域基因修复 (Global Gene Repair)
针对 WPS 生成的 Strict OOXML 格式文件或含有损坏外部链接的文件，本系统内置了强大的预处理清洗器 (`src/loader.py` & `src/preprocessor.py`)，确保 100% 可读取：
- **物理移除**：自动剔除导致崩溃的 `calcChain.xml`, `externalLinks`, `printerSettings`, `vbaProject.bin`。
- **命名空间洗白**：将 Strict OOXML 命名空间全局替换为 Microsoft 标准 Transitional OOXML。
- **外科手术式修复**：使用正则精准移除 `workbook.xml` 中导致崩溃的 `<definedNames>` 和 `<externalReferences>` 节点。
- **智能降级加载**：优先尝试完整加载，若失败则自动切换至 `read_only=True` 模式保底。

### 2. 多模式解析
- **标准表格**：自动处理合并单元格与多级表头展平。
- **语义锚点 (`parse_table_at_anchor`)**：通过关键词定位混合 Sheet 中的特定子表。
- **表单模式 (`parse_form`)**：智能提取非结构化的 "Key: Value" 对。
- **层级模式 (`parse_hierarchy`)**：基于缩进自动构建财务报表的树形 JSON 结构。

### 3. LLM 语义提取 (Qwen Integration)
- **上下文序列化 (`src/serializer.py`)**：将 Excel 数据转换为 Token 优化的 Markdown 格式，自动截断超长文本，过滤空行。
- **Prompt 工程 (`src/llm_bridge.py`)**：内置针对房地产/金融领域的 Prompt 模板，可直接对接 Qwen/GPT 进行结构化数据提取。

## 安装与设置

1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 生成测试数据：
   ```bash
   python scripts/generate_sample_excel.py  # 基础 V1 数据
   python scripts/generate_v2_data.py       # 高级 V2 数据（包含锚点、表单、层级）
   ```

## 使用示例

### 基础解析
```python
from src.parser import ExcelParser
import json

# 自动触发“全域基因修复”
parser = ExcelParser("data/complex_v2.xlsx")

# 1. 语义锚点定位
df_funding = parser.parse_table_at_anchor("Mixed Data", "Funding Assumptions")

# 2. 表单提取
project_info = parser.parse_form("Project Form")

# 3. 层级提取
financials = parser.parse_hierarchy("Financial Statement")
print(json.dumps(financials, indent=2, ensure_ascii=False))
```

### LLM 提取 Demo
```bash
python scripts/qwen_extraction_demo.py
```
该脚本演示了从 加载 -> 修复 -> 序列化 -> Prompt 构建 -> 模拟提取 的全流程。

## 工具脚本

- `scripts/analyze_excel.py`：**增强版分析工具**。包含完整的预处理逻辑，可用于诊断严重损坏的 Excel 文件结构。
  ```bash
  python scripts/analyze_excel.py <path_to_corrupt_file.xlsx>
  ```

## 测试

使用 pytest 运行全套测试（含集成测试）：
```bash
python -m pytest tests/
```
