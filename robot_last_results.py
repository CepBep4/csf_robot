"""Хранение результатов последнего запуска робота (mode=set) для выгрузки в xlsx."""
from datetime import datetime

# Результаты последнего запуска «Заполнить данные 1С»: список dict с ключами number_case, ok, message
last_results = None
last_mode = None
last_timestamp = None


def set_last_run(mode: str, results: list):
    global last_results, last_mode, last_timestamp
    last_results = results if results is not None else []
    last_mode = mode
    last_timestamp = datetime.now()
