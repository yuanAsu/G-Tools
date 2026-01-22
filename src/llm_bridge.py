from src.loader import load_complex_excel
from src.serializer import serialize_excel_to_markdown
import json

SYSTEM_PROMPT = """你是一个专业的房地产投资财务分析师。你的任务是从非结构化的 Excel 文本数据中提取关键的投资指标。"""

# Double curly braces for JSON example to escape them for .format()
USER_PROMPT_TEMPLATE = """以下是 Excel 文件的原始数据扫描结果：

Plaintext

{markdown_data}

请分析上述数据，精准提取以下字段，并以严格的 JSON 格式输出。如果找不到对应字段，请填 null，不要编造数据。

需要提取的字段：

project_name (项目名称)
city (城市)
total_parking_spaces (总车位数，请自动累加地下和地面车位)
cooperation_mode (合作模式)
project_duration_years (项目周期-年)
financials: {{ irr_5_year: (5年IRR，保留4位小数), payback_period: (投资回收期), first_year_profit: (首年利润，统一单位为万元) }}

Output JSON:"""

def generate_analysis_prompt(filepath: str) -> str:
    """
    Loads Excel, serializes it, and constructs the full prompt for the LLM.
    """
    wb = load_complex_excel(filepath)
    markdown_data = serialize_excel_to_markdown(wb)

    full_prompt = USER_PROMPT_TEMPLATE.format(markdown_data=markdown_data)
    return full_prompt

def simulate_qwen_extraction(prompt: str) -> dict:
    """
    Mocks the Qwen API call.
    In a real app, this would use `openai` client or similar to call Qwen model.
    """
    # This is a dummy response based on the expected data in complex_v2.xlsx
    mock_response = {
        "project_name": "Apollo Mission",
        "city": None,
        "total_parking_spaces": None,
        "cooperation_mode": None,
        "project_duration_years": None,
        "financials": {
            "irr_5_year": None,
            "payback_period": None,
            "first_year_profit": None
        }
    }
    return mock_response
