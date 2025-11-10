import json
import os
from pathlib import Path

from EI_protocols_utils.utils.settings import settings

weather_file = Path(settings.weather_journal_path).resolve()

def load_data():
    """Загружает данные из JSON-файла или возвращает пустой словарь, если файла нет."""
    if not os.path.exists(weather_file):
        return {}
    with open(weather_file, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    """Сохраняет данные в JSON-файл."""
    with open(weather_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
def add_weather(date, temperature, pressure, humidity):
    """
    Добавляет или обновляет данные о погоде для указанной даты.
    Args:
        date (str): Дата в формате "ДД.ММ.ГГГГ".
        temperature (str): Температура в градусах цельсия.
        pressure (str): Давление в КПа.
        humidity (str): Относительная влажность в %.
    """
    data = load_data()
    data[date] = {
        "temperature": temperature,
        "pressure": pressure,
        "humidity": humidity
    }
    save_data(data)
    
def get_weather(date):
    """
    Возвращает данные о погоде для указанной даты.
    """
    data = load_data()
    if date in data:
        return data[date]
    else:
        return None