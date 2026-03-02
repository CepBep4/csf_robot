# Загружаем xlsx с отчётом
from os import name
from turtle import position, right
from unittest import result
from pyautogui import click, press, rightClick, write, pixel, doubleClick, hotkey
from pyautogui import position as p
from utils import load_from_xlsx, addToBuffer, getFromBuffer, selectAll, paste
from time import sleep
data = load_from_xlsx("27022026.xlsx")
# 18-027367

"ПОИСК ДЕЛА"
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
        
        return True
    else:
        press('esc')
        sleep(2+cooldown)
        press('esc')
        sleep(2+cooldown)
        press('esc')
        sleep(2+cooldown)
        return False


def handle_court():
    "Заполнение и проверка вкладки СУД"
    cooldown = 0
    name_defedant = 'Козулин Владислав Валерьевич'
    # court = 'Свердловский районный суд г. Красноярска'
    court = 'wbfejbfehbfjerbj'
    date_base = "21.10.2024"
    date_plus_mounth = "05.09.2024"
    result_case = "удч" #удп - полностью удч - частично - отк - отказ в иске
    summ_requests_s = 400000
    summ_real_s = 400000
    summ_requests_g = 7200
    summ_real_g = 7200
    summ_s_check = False
    summ_g_check = False


    "Проваливаемся во вкладку суд"
    click(2518,630)
    sleep(3+cooldown)

    "- ПОЛЕ ОТВЕТЧИК -"
    "Нажимаем на строку ответчик"
    click(520,367)
    sleep(2+cooldown)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    "Проверяем на заполнение поле"
    if getFromBuffer() == 'none':
        "Нажимаем на кнопку чтобы переключить в режим строки"
        click(2502,369)
        sleep(2+cooldown)

        "В открывшемся меню переключаемся на строку"
        press("up")
        sleep(2+cooldown)
        press("enter")
        sleep(2+cooldown)

        "Вводим имя ответчика"
        addToBuffer(name_defedant)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

    "- ПОЛЕ СУД 1 - "

    "Нажимаем на поле СУД"
    click(660,416)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer() == 'none':
        "Нажимаем на поле СУД"
        click(660,416)
        
        "Вводим суд"
        addToBuffer(court)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

        "Проверяем что суд нашёлся в системе"
        if pixel(498,431) == (83, 106, 194):
            
            "Выбираем найденный суд"
            click(498,431)
            sleep(2+cooldown)
        
        else:
            "Возвращаем ошибку, прерываем заполнение"
            for _ in range(7):
                press('esc')
                sleep(2+cooldown)
                
            press('right')
            sleep(2+cooldown)
            press('enter')
            sleep(2+cooldown)
            
            for _ in range(5):
                press('esc')
                sleep(2+cooldown)
            return "Суд не найден"
            
    "- ПОЛЕ СУД 2 - "

    "Нажимаем на поле СУД"
    click(589, 466)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer() == 'none':
        "Нажимаем на поле СУД"
        click(589, 466)
        
        "Вводим суд"
        addToBuffer(court)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

        "Проверяем что суд нашёлся в системе"
        if pixel(422,480) == (83, 106, 194):
            
            "Выбираем найденный суд"
            click(422,480)
            sleep(2+cooldown)
        
        else:
            "Возвращаем ошибку, прерываем заполнение"
            for _ in range(7):
                press('esc')
                sleep(2+cooldown)
                
            press('right')
            sleep(2+cooldown)
            press('enter')
            sleep(2+cooldown)
            
            for _ in range(5):
                press('esc')
                sleep(2+cooldown)
                
            return "Суд не найден"

    "- ПОЛЕ ДАТА 1 -"
    "Нажимаем на дату 1"
    click(412,539)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer().strip() == '  .  .    '.strip():
        "Нажимаем на дату 1"
        click(412,539)
        
        "Выделяем всё, чтобы удалить"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)
        
        "Удаляем"
        press('backspace')
        sleep(2+cooldown)
        
        "Вводим дату 1"
        addToBuffer(date_base)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
    "- ПОЛЕ ДАТА 2 -"
    "Нажимаем на дату 2"
    click(1011,538)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer().strip() == '  .  .    '.strip():
        "Нажимаем на дату 2"
        click(1011,538)
        
        "Выделяем всё, чтобы удалить"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)
        
        "Удаляем"
        press('backspace')
        sleep(2+cooldown)
        
        "Вводим дату 2"
        addToBuffer(date_plus_mounth)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

    "- ПОЛЕ РЕЗУЛЬТАТ ПО ДЕЛУ -"

    "Нажимаем на выбор результата"
    click(2174, 538)
    sleep(2+cooldown)

    "Пролистываем в начало"
    for _ in range(12):
        press('up')
        sleep(1+cooldown)

    if result_case == "удп": #9
        for _ in range(9):
            press('down')
            sleep(1+cooldown)

    elif result_case == "удч": #10
        for _ in range(10):
            press('down')
            sleep(1+cooldown)


    elif result_case == "отк": #7
        for _ in range(7):
            press('down')
            sleep(1+cooldown)
            
    "Нажимаем enter чтобы применить"
    press("enter")
    sleep(2+cooldown)

    "ВЫПЛАТЫ"
    "Нажимаем на таб выплаты"
    click(269,295)

    "Нажимаем на первую строку на номер"
    click(197,363)
    sleep(1+cooldown)

    "Поверяем присутствие платежа"
    if pixel(197,363) != (85, 105, 194):
        "Платежей нет, добавляем первый"
        "Добавляем страховой платеж ЦБД000031"
        
        "Нажимаем на первую строку на номер"
        click(197,363)
        sleep(1+cooldown)
        
        "Создаём строку"
        press('down')
        sleep(2+cooldown)
        press('down')
        sleep(2+cooldown)
        
        "Переключаемся на поиск"
        hotkey('ctrl', 'f', interval=0.5)
        sleep(2+cooldown)
        
        "Вводим номер статуса платежа"
        addToBuffer("ЦБД000031")
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
        "Подтверждаем выбор"
        for _ in range(3):
            press('enter')
            sleep(2+cooldown)
            
        "Вводим номер статуса платежа"
        addToBuffer(summ_requests_s)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
        "Подтверждаем выбор"
        press('enter')
        sleep(2+cooldown)
        
        "Вводим номер статуса платежа"
        addToBuffer(summ_real_s)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
        "Подтверждаем выбор"
        press('enter')
        sleep(2+cooldown)
        
        "Гос пошлины ЦБД000016"
        
        "Создаём строку"
        press('down')
        sleep(2+cooldown)
        press('down')
        sleep(2+cooldown)
        
        "Переключаемся на поиск"
        hotkey('ctrl', 'f', interval=0.5)
        sleep(2+cooldown)
        
        "Вводим номер статуса платежа"
        addToBuffer("ЦБД000016")
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
        "Подтверждаем выбор"
        for _ in range(3):
            press('enter')
            sleep(2+cooldown)
            
        "Вводим номер статуса платежа"
        addToBuffer(summ_requests_g)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
        "Подтверждаем выбор"
        press('enter')
        sleep(2+cooldown)
        
        "Вводим номер статуса платежа"
        addToBuffer(summ_real_g)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
    else:  
        "Нажимаем на первую строку на номер"
        click(197,363)
        sleep(1+cooldown)
            
        "Переводим вправо, чтобы проверить что"
        press("right")
        sleep(1+cooldown)

        "Нажимаем enter, чтобы проверить название выплаты"
        press('enter')
        sleep(1+cooldown)

        "Выделяем всё, чтобы скопировать"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)

        "Копируем, чтобы проверить по буфферу"
        addToBuffer("none")
        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)

        "Проверяем чтобы существовала любая из выплат"
        if getFromBuffer() == "Страховое возмещение" or getFromBuffer() == "Государственная пошлина (истец)":
            if getFromBuffer() == "Страховое возмещение":
                summ_s_check=True
                "Проверяем запрошенную сумму"
                press('esc')
                sleep(1+cooldown)
                press('right')
                sleep(1+cooldown)
                press('enter')
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_requests_s)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                
                "Проверяем одобренную сумму"
                press('right')
                sleep(2+cooldown)
                press('enter')
                sleep(2+cooldown)
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_real_s)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                    
            if getFromBuffer() == "Государственная пошлина (истец)":
                summ_g_check=True
                "Проверяем запрошенную сумму"
                press('esc')
                sleep(1+cooldown)
                press('right')
                sleep(1+cooldown)
                press('enter')
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_requests_g)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                
                "Проверяем одобренную сумму"
                press('right')
                sleep(2+cooldown)
                press('enter')
                sleep(2+cooldown)
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_real_g)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                    
        "Нажимаем на первую строку на номер"
        click(197,363)
        sleep(1+cooldown)

        "Переводим вниз, чтобы проверить переключиться на следущую строку"
        press("down")
        sleep(1+cooldown)

        "Переводим вправо, чтобы проверить что"
        press("right")
        sleep(1+cooldown)

        "Нажимаем enter, чтобы проверить название выплаты"
        press('enter')
        sleep(1+cooldown)

        "Выделяем всё, чтобы скопировать"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)


        "Копируем, чтобы проверить по буфферу"
        addToBuffer("none")
        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)
        
        "Проверяем чтобы существовала любая из выплат"
        if getFromBuffer() == "Страховое возмещение" or getFromBuffer() == "Государственная пошлина (истец)":
            if getFromBuffer() == "Страховое возмещение":
                summ_s_check=True
                "Проверяем запрошенную сумму"
                press('esc')
                sleep(1+cooldown)
                press('right')
                sleep(1+cooldown)
                press('enter')
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_requests_s)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                
                "Проверяем одобренную сумму"
                press('right')
                sleep(2+cooldown)
                press('enter')
                sleep(2+cooldown)
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_real_s)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                    
            if getFromBuffer() == "Государственная пошлина (истец)":
                summ_g_check=True
                "Проверяем запрошенную сумму"
                press('esc')
                sleep(1+cooldown)
                press('right')
                sleep(1+cooldown)
                press('enter')
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_requests_g)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                
                "Проверяем одобренную сумму"
                press('right')
                sleep(2+cooldown)
                press('enter')
                sleep(2+cooldown)
                
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
                
                if getFromBuffer() == "0,00":
                    "Выделяем всё, чтобы заменить"
                    hotkey('ctrl', 'a', interval=0.5)
                    sleep(2+cooldown)
                    
                    "Вводим сумму"
                    addToBuffer(summ_real_g)
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(2+cooldown)
                    
                    press('enter')
                    sleep(2+cooldown)
                    
                else:
                    press('esc')
                    sleep(2+cooldown)
                    
        else:
            "Добавляем новую строку с выплатой"
            if summ_s_check == False:
                "Добавляем страховой платеж ЦБД000031"
                
                "Переключаемся на поиск"
                hotkey('ctrl', 'f', interval=0.5)
                sleep(2+cooldown)
                
                "Вводим номер статуса платежа"
                addToBuffer("ЦБД000031")
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
                "Подтверждаем выбор"
                for _ in range(3):
                    press('enter')
                    sleep(2+cooldown)
                    
                "Вводим номер статуса платежа"
                addToBuffer(summ_requests_s)
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
                "Подтверждаем выбор"
                press('enter')
                sleep(2+cooldown)
                
                "Вводим номер статуса платежа"
                addToBuffer(summ_real_s)
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
            
            if summ_g_check == False:
                "Гос пошлины ЦБД000016"
                
                "Переключаемся на поиск"
                hotkey('ctrl', 'f', interval=0.5)
                sleep(2+cooldown)
                
                "Вводим номер статуса платежа"
                addToBuffer("ЦБД000016")
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
                "Подтверждаем выбор"
                for _ in range(3):
                    press('enter')
                    sleep(2+cooldown)
                    
                "Вводим номер статуса платежа"
                addToBuffer(summ_requests_g)
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
                "Подтверждаем выбор"
                press('enter')
                sleep(2+cooldown)
                
                "Вводим номер статуса платежа"
                addToBuffer(summ_real_g)
                sleep(1+cooldown)
                hotkey('ctrl', 'v', interval=0.5)
                sleep(2+cooldown)
                
    "Нажимаем провести и закрыть"   
    click(246,265)   
    sleep(20+cooldown)
    
    "Дело передаём в суд"
    if pixel(2092,1343) == (204, 255, 216) and pixel(1952,1344) != (135, 206, 250):
        "Нажимаем на кнопку"
        click(2092,1343)
        sleep(25+cooldown)

    "Дело в пользу общества или нет"
    if pixel(2361,1342) == (146, 236, 146) and pixel(2232,1339) == (254, 163, 122):
    # print(pixel(2361,1342)) (146, 236, 146) зеленая
    # print(pixel(2232,1339)) (254, 163, 122) ораньжевая
        if result_case == "отк":
            "Нажимаем не в пользу общества"
            click(2232,1339)
            sleep(25+cooldown)
        else:
            "Нажимаем в пользу общества"
            click(2361,1342)
            sleep(25+cooldown)            
    
def ip_handler():
    cooldown=0
    view_ip_list="Оригинал исполнительного листа"
    ip_list_map = {
        "Оригинал исполнительного листа":3,
        "Постановление ФССП":4,
        "Инкассове списание из банка":2,
        "Дубликат испольнительного листа":1,
        "Добровольная оплата по решению суда":0  
    }

    number_ip_list="ФС №043214528"
    summ=407200.24
    data_get_ip_list="10.01.2025"
    
    "Открываем вкладу ИП"
    click(1687,774)
    sleep(2+cooldown)
    
    "Проверяем открылась ли вкладка"
    if pixel(1687,774) == (255,255,255):
        "Нажимаем на поле вид испольнительного листа"
        click(2511,369)
        sleep(2+cooldown)

        "Переключаемся на верхнюю позицию"
        for _ in range(6):
            press('up')
            sleep(1+cooldown)

        "Переключаемся на нужный вид листа"
        for _ in range(ip_list_map[view_ip_list]):
            press('down')
            sleep(1+cooldown)
            
        "Подтверждаем выбор"
        press('enter')
        sleep(2+cooldown)
        
        "Нажимаем на поле номер исполнительного листа"
        click(572,414)
        sleep(2+cooldown)

        "Выделяем всё, чтобы скопировать"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)

        "Копируем, чтобы проверить по буфферу"
        addToBuffer("none")
        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)

        if getFromBuffer() == 'none':
            "Нажимаем на поле номер испольнительного листа"
            click(572,414)
            sleep(2+cooldown)
            
            "Вводим номер исполнительного листа"
            addToBuffer(number_ip_list)
            sleep(1+cooldown)
            hotkey('ctrl', 'v', interval=0.5)
            sleep(2+cooldown)
            
        "Нажимаем на поле сумма"
        click(518,464)
        sleep(2+cooldown)

        "Выделяем всё, чтобы скопировать"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)

        "Копируем, чтобы проверить по буфферу"
        addToBuffer("none")
        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)

        if getFromBuffer() == '0,00':
            "Нажимаем на поле сумма"
            click(518,464)
            sleep(2+cooldown)
            
            "Вводим сумму"
            addToBuffer(summ)
            sleep(1+cooldown)
            hotkey('ctrl', 'v', interval=0.5)
            sleep(2+cooldown)
            
        "- ПОЛЕ ДАТА ip -"
        "Нажимаем на дату ip"
        click(2504,416)

        "Выделяем всё, чтобы скопировать"
        hotkey('ctrl', 'a', interval=0.5)
        sleep(2+cooldown)

        "Копируем, чтобы проверить по буфферу"
        addToBuffer("none")
        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)

        if getFromBuffer().strip() == '  .  .    '.strip():
            "Нажимаем на дату ip"
            click(2504,416)
            
            "Выделяем всё, чтобы удалить"
            hotkey('ctrl', 'a', interval=0.5)
            sleep(2+cooldown)
            
            "Удаляем"
            press('backspace')
            sleep(2+cooldown)
            
            "Вводим дату 2"
            addToBuffer(data_get_ip_list)
            sleep(1+cooldown)
            hotkey('ctrl', 'v', interval=0.5)
            sleep(2+cooldown)
            
        "Нажимаем провести и закрыть"
        click(356,266)
        sleep(2+cooldown)
        
        "Передаём в ФССП"
        if pixel(2159,1344) == (145, 239, 145) and pixel(1952,1344) == (135, 206, 250):
            "Нажимаем на кнопку"
            click(2159,1344)
            sleep(25+cooldown)
    else:
        "Закрываем вкладку"
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)   
        return "Вкладка ИП не открылась"
        
def download_cases(number_case, cooldown=0):
    "Нажимаем на скачать дело"
    click(273,268)
    sleep(180+cooldown)
    
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
    
    
# search_case("37-004741")
# # sleep(10)
# handle_court()
# # sleep(10)
# ip_handler()

for i in ['69-004134','03-003693','08-003703']:
    search_case(i)
    download_cases(i)
    sleep(5)








