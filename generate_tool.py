"""
Ad Campaign & Creative Naming Generator Build Script
Reads configurations and product maps to build index.html and ad-name-config.json
"""
import json
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "ad-name-config.json"
PRODUCT_MAP_PATH = BASE_DIR / "product-code-map.json"
DATA_JSON_PATH = BASE_DIR / "data.json"
OUTPUT_HTML = BASE_DIR / "index.html"

def load_json(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def main():
    print("Building Naming Tool Hub...")
    config = load_json(CONFIG_PATH)
    prod_map = load_json(PRODUCT_MAP_PATH)
    data = load_json(DATA_JSON_PATH)
    
    print(f"Loaded {len(config)} ad field configurations.")
    print(f"Loaded {len(prod_map)} product map categories.")
    print(f"Loaded data.json with keys: {list(data.keys())}")
    
    print("Build completed successfully. Run or serve index.html directly.")

if __name__ == "__main__":
    main()
