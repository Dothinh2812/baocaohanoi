import json
from pathlib import Path

recipes_dir = Path("recipes")
results = {}

for recipe_path in recipes_dir.glob("*.json"):
    try:
        data = json.loads(recipe_path.read_text(encoding="utf-8"))
        name = data.get("name", recipe_path.stem)
        results[name] = {
            "report_id": data.get("report_id", ""),
            "report_page_url": data.get("report_page_url", ""),
            "input_params": [p.get("name") for p in data.get("export_payload", {}).get("lstInputParams", [])]
        }
    except Exception as e:
        print(f"Error reading {recipe_path}: {e}")

print(json.dumps(results, indent=2, ensure_ascii=False))
