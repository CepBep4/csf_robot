# Загружаем xlsx с отчётом
from os import name
from turtle import position, right
from unittest import result
from pyautogui import click, press, rightClick, write, pixel, doubleClick, hotkey
from pyautogui import position as p
from worker.utils import load_from_xlsx, addToBuffer, getFromBuffer, selectAll, paste
from time import sleep

#Поиск дела открытие вкладки с поиском
def search_case(number_case, cooldown = 0):
    "#Нажимаем мои задачи"
    click(56,160)
    sleep(15+cooldown)

    "#Открываем поиск"
    rightClick(822,337)
    sleep(3+cooldown)

    "#Нажимаем найти лупу"
    press("down")
    sleep(3+cooldown)
    
    "Нажимаем поиск ентер"
    press('enter')
    sleep(2+cooldown)
    
    "Вводим имя ответчика"
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
    
    if pixel(479,386) == (255, 255, 255):
        "Нажимаем поиск ентер"
        press('enter')
        sleep(2+cooldown)
        
        return (f"Дело {number_case} успешно найдено", True)
    else:
        press('esc')
        sleep(2+cooldown)
        press('esc')
        sleep(2+cooldown)
        press('esc')
        sleep(2+cooldown)
        return (f"Дело {number_case} не обнаружено", False)


search_case('18-026543', 0)




