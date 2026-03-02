"""Общее состояние робота: флаг запроса остановки для прерывания длительной работы."""
import threading

_stop_requested = threading.Event()


def clear_stop():
    _stop_requested.clear()


def request_stop():
    _stop_requested.set()


def is_stop_requested() -> bool:
    return _stop_requested.is_set()


def request_stop_run():
    """Вызвать извне (например, с сервера) для запроса остановки текущего запуска."""
    _stop_requested.set()
