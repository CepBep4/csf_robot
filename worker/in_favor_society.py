from time import sleep

from pyautogui import click, press
from worker import BUTTON_PANEL_REGION
from worker.court_tab import (
    GREEN_BUTTON_RGB,
    RED_BUTTON_RGB,
    find_button_center_in_region,
)


def press_in_favor_society(result_case: str, cooldown: float = 0):
    """
    Прожимает кнопку "В пользу общества" или "Не в пользу общества"
    в зависимости от result_case.
    """
    target_rgb = RED_BUTTON_RGB if str(result_case).strip().lower() == "отк" else GREEN_BUTTON_RGB
    pos = find_button_center_in_region(BUTTON_PANEL_REGION, target_rgb, tolerance_abs=15, min_pixels=10)
    pos_ip = find_button_center_in_region(BUTTON_PANEL_REGION, (135, 206, 250), tolerance_abs=15, min_pixels=10)
    if not pos or pos_ip:
        return ("Кнопка 'В пользу/Не в пользу общества' не обнаружена", False)

    click(pos[0], pos[1])
    sleep(30 + cooldown)
    press("esc")
    sleep(10 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    return ("Кнопка 'В пользу/Не в пользу общества' успешно нажата", True)
