import os
import sys
import json
from src.llm_bridge import generate_analysis_prompt, simulate_qwen_extraction

def main():
    # Ensure data exists
    data_file = "data/complex_v2.xlsx"
    if not os.path.exists(data_file):
        print(f"Data file {data_file} not found. Running generator...")
        # Assuming we are in root
        if os.path.exists("scripts/generate_v2_data.py"):
            os.system("python scripts/generate_v2_data.py")
        else:
            print("Generator script not found.")
            sys.exit(1)

    print(f"--- Processing {data_file} ---")

    try:
        # 1. Generate Prompt (Gatekeeper Load -> Serialize -> Template)
        prompt = generate_analysis_prompt(data_file)

        print("\n[Generated Prompt Preview (First 500 chars)]:")
        print("-" * 40)
        print(prompt[:500] + "...")
        print("-" * 40)

        # 2. Simulate Extraction
        print("\n[Simulating Qwen Extraction]...")
        result = simulate_qwen_extraction(prompt)

        print("\n[Extracted JSON Result]:")
        print(json.dumps(result, indent=2, ensure_ascii=False))

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
