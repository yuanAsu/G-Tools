# [中文](README.md) | [English](README_EN.md)

# 复杂 Excel 解析器 (Complex Excel Parser)

这是一个用于解析复杂结构 Excel 文件的 Python 应用程序。它支持合并单元格处理、多级表头展平、语义锚点定位、键值对表单提取以及层级数据提取。

## 功能特性

- **标准表格解析：**
  - 自动处理合并单元格（将值填充到整个合并区域）。
  - 将多级表头合并为单行表头（例如 "Q1 - Sales"）。

- **语义锚点定位 (`parse_table_at_anchor`)：**
  - 通过关键词（如“融资假设”）在混合内容的 Sheet 中定位特定表格。
  - 提取相对于锚点的数据区域。

- **表单提取 (`parse_form`)：**
  - 扫描 Sheet 中的键值对（Key-Value），例如 "项目名称: Apollo"。
  - 适用于提取项目详情、封面信息等非结构化布局。

- **层级提取 (`parse_hierarchy`)：**
  - 根据缩进列构建嵌套的 JSON/字典结构。
  - 非常适合处理财务报表（如 资产 -> 流动资产 -> 现金）。

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

```python
from src.parser import ExcelParser
import json

parser = ExcelParser("data/complex_v2.xlsx")

# 1. 语义锚点定位
df_funding = parser.parse_table_at_anchor("Mixed Data", "Funding Assumptions")
print(df_funding)

# 2. 表单提取
project_info = parser.parse_form("Project Form")
print(project_info)
# 输出: {'Project Name': 'Apollo Mission', ...}

# 3. 层级提取
financials = parser.parse_hierarchy("Financial Statement")
print(json.dumps(financials, indent=2, ensure_ascii=False))
```

## 工具脚本

- `scripts/analyze_excel.py`：检查 Excel 文件的结构（Sheet 名称、合并单元格情况等）。
  ```bash
  python scripts/analyze_excel.py data/complex_v2.xlsx
  ```

## 测试

使用 pytest 运行测试：
```bash
python -m pytest tests/
```
