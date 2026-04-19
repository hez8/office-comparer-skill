import json
import os
import argparse
import sys

def prepare_web(file_a, file_b):
    # 确保是绝对路径
    abs_a = os.path.abspath(file_a)
    abs_b = os.path.abspath(file_b)
    
    config = {
        "file_a": abs_a,
        "file_b": abs_b,
        "timestamp": os.path.getmtime(abs_a) if os.path.exists(abs_a) else 0
    }
    
    # 配置文件路径：与此脚本同目录
    config_path = os.path.join(os.path.dirname(__file__), "auto_load.json")
    
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"Web configuration updated: {config_path}")
        return True
    except Exception as e:
        print(f"Failed to update Web configuration: {e}", file=sys.stderr)
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Prepare Office-Comparer Web preview.")
    parser.add_argument("--file_a", required=True, help="Path to first file")
    parser.add_argument("--file_b", required=True, help="Path to second file")
    
    args = parser.parse_args()
    prepare_web(args.file_a, args.file_b)
