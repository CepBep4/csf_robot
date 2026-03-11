"""Запуск робота: сброс флага остановки и логирование в robot.log."""
from robot import interface_start
from robot_log import get_logger
from robot_state import clear_stop
from robot_last_results import set_last_run


def run(mode: str, data: dict | list):
    """Запуск interface_start с подготовкой (clear_stop) и логированием в robot.log."""
    clear_stop()
    log = get_logger()
    log.info("Старт: mode=%s", mode)
    try:
        out = interface_start(mode, data)
        if mode == "set" and out is not None:
            set_last_run(mode, out)
        log.info("Завершено: mode=%s", mode)
    except Exception as e:
        log.exception("Ошибка при выполнении: %s", e)
        raise
