from pyautogui import click, press, rightClick, pixel, hotkey
from time import sleep
from worker.utils import addToBuffer, getFromBuffer

#Поиск дела открытие вкладки с поиском
def search_case(number_case, cooldown = 0):
    "#Нажимаем мои задачи"
    hotkey('alt', '2', interval=0.5)
    sleep(1+cooldown)
    for _ in range(2):
        press('down')
        sleep(1+cooldown)
    press('up')
    sleep(1+cooldown)
    press('enter')
    sleep(15+cooldown)
    
    "Переходим на поиск"
    for _ in range(3):
        press('tab')
        sleep(0.5+cooldown)
        
    "Передвигаемся на нужную колонку"
    for _ in range(2):
        press('right')
        sleep(0.5+cooldown)   

    "#Открываем поиск"
    hotkey('ctrl', 'f', interval=0.5)
    sleep(3+cooldown)
    
    "Вводим номер дела"
    addToBuffer(number_case)
    sleep(1+cooldown)
    hotkey('ctrl', 'v', interval=0.5)
    sleep(2+cooldown)
    
    "Нажимаем поиск ентер"
    press('enter')
    sleep(2+cooldown)
    
    "Нажимаем поиск ентер"
    press('enter')
    sleep(2+cooldown)
    
    "Проверяем что нашли"
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)
    if number_case in getFromBuffer():
        "Нажимаем поиск ентер"
        press('enter')
        sleep(2+cooldown)
        
        return (f"Дело {number_case} успешно найдено", True)
    else:
        press('esc')
        sleep(1+cooldown)
        press('esc')
        sleep(1+cooldown)
        press('esc')
        sleep(1+cooldown)
        return (f"Дело {number_case} не обнаружено", False)