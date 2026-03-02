from pyautogui import click, press, pixel, hotkey
from time import sleep
from worker.utils import addToBuffer, getFromBuffer

def court_tab(
    name_defedant: str,  # Имя ответчика
    court: str,  # Суд (название)
    date_base: str,  # Дата вынесения решения, формат дд.мм.гггг
    date_plus_mounth: str,  # Дата вступления решения в силу, формат дд.мм.гггг
    result_case: str,  # удп | удч | отк
    summ_requests_s: float,  # Запрошенная сумма (страховое возмещение)
    summ_real_s: float,  # Одобренная сумма (страховое возмещение)
    summ_requests_g: float,  # Запрошенная сумма (госпошлина)
    summ_real_g: float,  # Одобренная сумма (госпошлина)
    cooldown=0,
):   
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
            return ("Суд не найден", False)
            
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
                
            return ("Суд не найден", False)

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
    
    _case_to_court = False
    "Дело передаём в суд"
    if pixel(2092,1343) == (204, 255, 216) and pixel(1952,1344) != (135, 206, 250):
        _case_to_court = True
        "Нажимаем на кнопку"
        click(2092,1343)
        sleep(25+cooldown)

    _case_to_public = False
    "Дело в пользу общества или нет"
    if pixel(2361,1342) == (146, 236, 146) and pixel(2232,1339) == (254, 163, 122):
        _case_to_public = True
        if result_case == "отк":
            "Нажимаем не в пользу общества"
            click(2232,1339)
            sleep(25+cooldown)
        else:
            "Нажимаем в пользу общества"
            click(2361,1342)
            sleep(25+cooldown)
            
    return (f"Вкладка суд успешно заполнена. Прожата кнопка передать дело в суд: {_case_to_court}. Прожата кнопка В пользу/Не в пользу общества: {_case_to_public}.", True)       
   
def court_tab_check(cooldown=0):   
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
    name_defedant = getFromBuffer()

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
    court = getFromBuffer()

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
    date_base = getFromBuffer()
        
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
    date_plus_mounth = getFromBuffer()

    "ВЫПЛАТЫ"
    "Нажимаем на таб выплаты"
    click(269,295)

    "Нажимаем на первую строку на номер"
    click(197,363)
    sleep(1+cooldown)

    "Поверяем присутствие платежа"
    if pixel(197,363) != (85, 105, 194):
        summ_s_get = "none"
        summ_s_req = "none"
        summ_g_get = "none"
        summ_g_req = "none"
    else:
        summ_s_get = "none"
        summ_s_req = "none"
        summ_g_get = "none"
        summ_g_req = "none"
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
                summ_s_get = getFromBuffer()
                
                "Подтверждаем"
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
                summ_s_req = getFromBuffer()
                
                "Подтверждаем"
                press('esc')
                sleep(2+cooldown)                
                
            if getFromBuffer() == "Государственная пошлина (истец)":
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
                summ_g_get = getFromBuffer()
                
                "Подтверждаем"
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
                summ_g_req = getFromBuffer()
                
                "Подтверждаем"
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
                summ_s_get = getFromBuffer()
                
                "Подтверждаем"
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
                summ_s_req = getFromBuffer()
                
                "Подтверждаем"
                press('esc')
                sleep(2+cooldown)
                    
            if getFromBuffer() == "Государственная пошлина (истец)":
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
                summ_g_get = getFromBuffer()
                
                "Подтверждаем"
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
                summ_g_req = getFromBuffer()
                
                "Подтверждаем"
                press('esc')
                sleep(2+cooldown)
                
    "Закрываем вкладку"   
    press('esc')  
    sleep(20+cooldown)
            
    return (f"Вкладка суд успешно проверена", True, {
        'name_defedant': name_defedant,
        'court': court,
        'date_base': date_base,
        'date_plus_mounth': date_plus_mounth,
        'summ_s_get': summ_s_get,
        'summ_s_req': summ_s_req,
        'summ_g_get': summ_g_get,
        'summ_g_req': summ_g_req,
    })       
   
