from time import sleep

from pyautogui import click, press
from worker import BUTTON_PANEL_REGION
from worker.ip_tab import (
    GREEN_IP_RGB,
    find_button_center_in_region,
)


def press_sent_to_fssp(cooldown: float = 0):
    """
    Прожимает кнопку "ИП получено / Передано в ФССП".
    Логика вынесена из ip_tab в отдельный шаг pipeline.
    """
    pos = find_button_center_in_region(
        BUTTON_PANEL_REGION,
        GREEN_IP_RGB,
        tolerance_abs=15,
        min_pixels=10,
    )
    if not pos:
        return ("Кнопка 'Передано в ФССП' не обнаружена", False)

    click(pos[0], pos[1])
    for _ in range(3):
        press('esc')
        sleep(10+cooldown)
    return ("Кнопка 'Передано в ФССП' успешно нажата", True)
