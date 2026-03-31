from time import sleep

from pyautogui import click, press
from worker import BUTTON_PANEL_REGION
from worker.court_tab import (
    GREEN_TRANSFER_RGB,
    find_button_center_in_region,
)


def press_case_to_court(cooldown: float = 0):
    """
    Прожимает кнопку "Передать дело в суд".
    Логика вынесена из court_tab в отдельный модуль.
    """
    pos = find_button_center_in_region(
        BUTTON_PANEL_REGION,
        GREEN_TRANSFER_RGB,
        tolerance_abs=15,
        min_pixels=10,
    )
    if not pos:
        return ("Кнопка 'Передать дело в суд' не обнаружена", False)

    click(pos[0], pos[1])
    sleep(30 + cooldown)
    press("esc")
    sleep(10 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    return ("Кнопка 'Передать дело в суд' успешно нажата", True)
