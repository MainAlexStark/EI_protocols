import json

def save_paths(paths, filename):
    """Сохраняет словарь путей в JSON-файл."""
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(paths, f, ensure_ascii=False, indent=4)

def load_paths(filename):
    """Загружает словарь путей из JSON-файла."""
    with open(filename, "r", encoding="utf-8") as f:
        return json.load(f)