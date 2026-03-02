from pyautogui import click, press, hotkey
from time import sleep
from worker.utils import addToBuffer

def download_mode(number_case, cooldown=0, cooldown_download = 180):
    "Нажимаем на скачать дело"
    click(273,268)
    sleep(cooldown_download+cooldown)
    
    "Выделяем всё, чтобы заменить"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)
    
    "Вводим имя дела"
    addToBuffer(number_case)
    sleep(1+cooldown)
    hotkey('ctrl', 'v', interval=0.5)
    sleep(2+cooldown) 
    
    "Нажимаем сохранить"
    press('enter')
    sleep(2+cooldown)
    
    "Закрываем сессию"
    press('esc')
    sleep(2+cooldown)
    press('esc')
    sleep(2+cooldown)