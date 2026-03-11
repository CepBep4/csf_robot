from pyautogui import click, press, pixel, hotkey, screenshot
from time import sleep
from worker.utils import addToBuffer, getFromBuffer

def check_color(target_rgb, tolerance=5, region=None):
    """
    Ищет на экране (или в области region) хотя бы один пиксель,
    цвет которого совпадает с target_rgb с допуском tolerance по каждому каналу.
    target_rgb: (r, g, b), например (85, 105, 194)
    region: None = весь экран, иначе (left, top, width, height)
    Возвращает True если найден, иначе False.
    """
    img = screenshot(region=region)
    w, h = img.size
    tr, tg, tb = target_rgb
    for x in range(w):
        for y in range(h):
            p = img.getpixel((x, y))
            r, g, b = p[:3]
            if abs(r - tr) <= tolerance and abs(g - tg) <= tolerance and abs(b - tb) <= tolerance:
                return True
    return False

def pixel_matches(x, y, target_rgb, tolerance_pct=0.05):
    """Проверяет, что цвет пикселя (x, y) совпадает с target_rgb с допуском tolerance_pct (0.05 = 5%)."""
    r, g, b = pixel(x, y)
    tr, tg, tb = target_rgb
    t = int(255 * tolerance_pct)
    return abs(r - tr) <= t and abs(g - tg) <= t and abs(b - tb) <= t

def _pixel_matches(img, px, py, target_rgb, tol, green_only=False):
    """Проверка пикселя: цвет в допуске tol (абсолютное значение по каналу). green_only=True — G > R."""
    if px < 0 or py < 0 or px >= img.size[0] or py >= img.size[1]:
        return False
    p = img.getpixel((px, py))
    r, g, b = p[0], p[1], p[2]
    tr, tg, tb = target_rgb
    if not (abs(r - tr) <= tol and abs(g - tg) <= tol and abs(b - tb) <= tol):
        return False
    if green_only:
        if g <= r:
            return False
    return True

def find_pixel_in_region(region, target_rgb, tolerance_pct=0.05):
    """
    Ищет в области region первый пиксель цвета target_rgb (с допуском 5%).
    region: (left, top, width, height)
    Возвращает (screen_x, screen_y) или None.
    """
    img = screenshot(region=region)
    w, h = img.size
    left, top = region[0], region[1]
    tr, tg, tb = target_rgb
    t = int(255 * tolerance_pct)
    for py in range(h):
        for px in range(w):
            p = img.getpixel((px, py))
            r, g, b = p[:3]
            if abs(r - tr) <= t and abs(g - tg) <= t and abs(b - tb) <= t:
                return (left + px, top + py)
    return None

def find_button_center_in_region(region, target_rgb, tolerance_pct=0.15, tolerance_abs=None, min_pixels=30, green_only=False):
    """
    Ищет в области первое пятно цвета target_rgb, возвращает центр пятна.
    tolerance_abs: если задан — допуск по каналу в единицах (например 15 = ±15 от значения); иначе tolerance_pct.
    """
    img = screenshot(region=region)
    w, h = img.size
    left, top = region[0], region[1]
    t = int(tolerance_abs) if tolerance_abs is not None else int(255 * tolerance_pct)

    def matches(px, py):
        return _pixel_matches(img, px, py, target_rgb, t, green_only=green_only)

    for py in range(h):
        for px in range(w):
            if not matches(px, py):
                continue
            # BFS — собрать связное пятно
            stack = [(px, py)]
            seen = {(px, py)}
            while stack:
                cx, cy = stack.pop()
                for dx, dy in ((0, 1), (1, 0), (0, -1), (-1, 0)):
                    nx, ny = cx + dx, cy + dy
                    if (nx, ny) not in seen and 0 <= nx < w and 0 <= ny < h and matches(nx, ny):
                        seen.add((nx, ny))
                        stack.append((nx, ny))
            if len(seen) < min_pixels:
                continue
            n = len(seen)
            cx = sum(p[0] for p in seen) // n
            cy = sum(p[1] for p in seen) // n
            return (left + cx, top + cy)
    return None

# Область панели кнопок: левый верх (1710, 1248), правый низ (2545, 1363)
BUTTON_PANEL_REGION = (1710, 1248, 835, 115)
# Зелёная «Передать в суд» (другой оттенок)
GREEN_TRANSFER_RGB = (204, 255, 216)
# Зелёная «В пользу общества», красная «Не в пользу общества»
GREEN_BUTTON_RGB = (146, 236, 146)
RED_BUTTON_RGB = (255, 160, 122)

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
    "Проваливаемся во вкладку суд"
    for _ in range(21):
        press('tab')
        sleep(0.5 + cooldown)
        
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
        press('esc')
        sleep(10+cooldown)
        return ("Суд ", False)        
        
    hotkey('ctrl', 'shift', 'f4', interval=0.5)
    sleep(3+cooldown)

    "- ПОЛЕ ОТВЕТЧИК -"
    "Нажимаем на строку ответчик"
    press('tab')
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
        for _ in range(2):
            press('up')
            sleep(0.5+cooldown)
        press('enter')
        sleep(2+cooldown)

        "Вводим имя ответчика"
        addToBuffer(name_defedant)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

    "- ПОЛЕ СУД 1 - "

    "Нажимаем на поле СУД"
    for _ in range(2):
        press('tab')
        sleep(0.5+cooldown)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer() == 'none':
        "Вводим суд"
        addToBuffer(court)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

        "Проверяем что суд нашёлся в системе"
        if not check_color((83, 106, 194)):
            "Возвращаем ошибку, прерываем заполнение"
            for _ in range(3):
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
            
    # "- ПОЛЕ СУД 2 - "

    "Нажимаем на поле СУД"
    press('enter')
    sleep(0.5+cooldown) 
    press('tab')
    sleep(0.5+cooldown)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)
    
    if getFromBuffer() == 'none':
        "Вводим суд"
        addToBuffer(court)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

        "Проверяем что суд нашёлся в системе"
        if not check_color((83, 106, 194)):
            "Возвращаем ошибку, прерываем заполнение"
            for _ in range(3):
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
    press('enter')
    sleep(0.5+cooldown)
    press('tab')
    sleep(0.5+cooldown)
    press('tab')
    sleep(0.5+cooldown)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer().strip() == '  .  .    '.strip():
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
    for _ in range(3):
        press('tab')
        sleep(0.5+cooldown)

    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer().strip() == '  .  .    '.strip():   
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
    for _ in range(2):
        press('tab')
        sleep(0.5+cooldown)

    "Пролистываем в начало"
    for _ in range(13):
        press('up')
        sleep(0.5+cooldown)

    if result_case == "удп": #9
        for _ in range(9):
            press('down')
            sleep(0.5+cooldown)

    elif result_case == "удч": #10
        for _ in range(10):
            press('down')
            sleep(0.5+cooldown)


    elif result_case == "отк": #7
        for _ in range(7):
            press('down')
            sleep(0.5+cooldown)
            
    "Нажимаем enter чтобы применить"
    press("enter")
    sleep(2+cooldown)

    "ВЫПЛАТЫ"
    "Нажимаем на таб выплаты"
    hotkey('ctrl', 'pagedown', interval=0.5)

    "Нажимаем на первую строку на номер"
    press('right')
    sleep(1+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)
    
    "Поверяем присутствие платежа"
    if getFromBuffer() == 'none':
        "Платежей нет"
        "Добавляем первый"
        press('down')
        sleep(0.5+cooldown)
        press('enter')
        sleep(3+cooldown)
        
        "Ищем страховой платеж"
        hotkey('ctrl','f',interval=0.5)

        "Вводим номер статуса платежа"
        addToBuffer("ЦБД000031")
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(10+cooldown)
        
        "Подтверждаем выбор"
        for _ in range(2):
            press('enter')
            sleep(2+cooldown)

        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)
        
        if getFromBuffer() == 'Страховое возмещение':
            press('enter')
            sleep(1+cooldown)
        
        else:
            "Завершаем текущую сессию"
            press('esc')
            sleep(1+cooldown)
            press('esc')
            sleep(1+cooldown)
            press('right')
            sleep(1+cooldown)
            press('enter')
            for _ in range(5):
                press('esc')
                sleep(1+cooldown)
            return ("В справочнике нет Страховое возмещение", False)

        "Вводим номер запрошенную сумму"
        addToBuffer(summ_requests_s)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(0.5+cooldown)
        
        sleep(0.5+cooldown)
        press('enter')
        
        "Вводим реальную сумму сумму"
        addToBuffer(summ_real_s)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(0.5+cooldown)
        press('enter')
        sleep(0.5+cooldown)
        
        "Открываем окно поиска"
        sleep(0.5+cooldown)
        press('enter')
        sleep(3+cooldown)
        
        "Ищем госпощлину платеж"
        hotkey('ctrl','f',interval=0.5)

        "Вводим номер статуса платежа"
        addToBuffer("ЦБД000016")
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(10+cooldown)
        
        "Подтверждаем выбор"
        for _ in range(2):
            press('enter')
            sleep(2+cooldown)

        sleep(1+cooldown)
        hotkey('ctrl', 'c', interval=0.5)
        sleep(1+cooldown)
        
        if getFromBuffer() == 'Государственная пошлина (истец)':
            press('enter')
            sleep(1+cooldown)
        
        else:
            "Завершаем текущую сессию"
            press('esc')
            sleep(1+cooldown)
            press('esc')
            sleep(1+cooldown)
            press('right')
            sleep(1+cooldown)
            press('enter')
            for _ in range(5):
                press('esc')
                sleep(1+cooldown)
            return ("В справочнике нет Госпошлины", False)
        
        "Вводим номер запрошенную сумму"
        addToBuffer(summ_requests_g)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(0.5+cooldown)
        
        sleep(0.5+cooldown)
        press('enter')
        
        "Вводим реальную сумму сумму"
        addToBuffer(summ_real_g)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(0.5+cooldown)
        
        "Кнопку провести и закрыть"
        for _ in range(8):
            hotkey('shift', 'tab', interval=0.5)
            sleep(0.1+cooldown)
        
    else:
        rcg = False
        rcs = False
        while not rcg or not rcs:
            while getFromBuffer() != "Государственная пошлина (истец)" and getFromBuffer() != "Страховое возмещение" and getFromBuffer() != 'none':
                press('down')
                "Копируем, чтобы проверить по буфферу"
                addToBuffer("none")
                sleep(1+cooldown)
                hotkey('ctrl', 'c', interval=0.5)
                sleep(1+cooldown)
            else:
                buffer = getFromBuffer()
                if buffer == 'Государственная пошлина (истец)' or buffer == "Страховое возмещение":
                    press('right')
                    sleep(0.5+cooldown)
                    
                    "Копируем, чтобы проверить по буфферу"
                    addToBuffer("none")
                    sleep(0.5+cooldown)
                    hotkey('ctrl', 'c', interval=0.5)
                    sleep(0.5+cooldown)
                    
                    "Проверяем чтобы сумма была указана иначе заполняем"
                    if getFromBuffer() == 'none':
                        press('enter')
                        sleep(0.5+cooldown)
                        "Вводим, запрошенную сумму"
                        addToBuffer(summ_requests_g if buffer == "Государственная пошлина (истец)" else summ_requests_s)
                        sleep(0.5+cooldown)
                        hotkey('ctrl', 'v', interval=0.5)
                        sleep(0.5+cooldown)  
                        press('enter')
                        sleep(0.5+cooldown)  
                        
                    press('right')                
                    sleep(0.5+cooldown) 
                    
                    "Копируем, чтобы проверить по буфферу"
                    addToBuffer("none")
                    sleep(0.5+cooldown)
                    hotkey('ctrl', 'c', interval=0.5)
                    sleep(0.5+cooldown)
                    
                    "Проверяем чтобы сумма была указана иначе заполняем"
                    if getFromBuffer() == 'none':
                        press('enter')
                        sleep(0.5+cooldown)
                        "Вводим, запрошенную сумму"
                        addToBuffer(summ_real_g  if buffer == "Государственная пошлина (истец)" else summ_real_s)
                        sleep(0.5+cooldown)
                        hotkey('ctrl', 'v', interval=0.5)
                        sleep(0.5+cooldown)  
                        press('enter')
                        sleep(0.5+cooldown) 
                    
                    press('left')
                    sleep(0.5+cooldown) 
                    press('left')
                    sleep(0.5+cooldown) 
                    if buffer == "Государственная пошлина (истец)":
                        rcg=True
                    else:
                        rcs = True
                        
                if buffer == 'none':
                    choice = 'none'
                    press('enter')
                    sleep(3+cooldown)
                    
                    "Ищем страховой платеж"
                    hotkey('ctrl','f',interval=0.5)

                    "Вводим номер статуса платежа"
                    if rcg == False:
                        choice = 'rcg'
                        addToBuffer("ЦБД000016")
                    elif rcs == False:
                        choice = 'rcs'
                        addToBuffer("ЦБД000031")
                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(10+cooldown)
                    
                    "Подтверждаем выбор"
                    for _ in range(2):
                        press('enter')
                        sleep(2+cooldown)

                    sleep(1+cooldown)
                    hotkey('ctrl', 'c', interval=0.5)
                    sleep(1+cooldown)
                    
                    if (getFromBuffer() == 'Страховое возмещение' and choice == 'rcs') or (getFromBuffer() == "Государственная пошлина (истец)" and choice == 'rcg'):
                        press('enter')
                        sleep(1+cooldown)
                    
                    else:
                        "Завершаем текущую сессию"
                        press('esc')
                        sleep(1+cooldown)
                        press('esc')
                        sleep(1+cooldown)
                        press('right')
                        sleep(1+cooldown)
                        press('enter')
                        for _ in range(5):
                            press('esc')
                            sleep(1+cooldown)
                        return ("В справочнике нет Страховое возмещение", False)

                    "Вводим номер запрошенную сумму"
                    if choice == 'rcs':
                        addToBuffer(summ_requests_s)
                    elif choice == 'rcg':
                        addToBuffer(summ_requests_g)

                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(0.5+cooldown)
                    
                    sleep(0.5+cooldown)
                    press('enter')
                    
                    "Вводим реальную сумму сумму"
                    if choice == 'rcs':
                        addToBuffer(summ_real_s)                  
                    elif choice == 'rcg':
                        addToBuffer(summ_real_g)

                    sleep(1+cooldown)
                    hotkey('ctrl', 'v', interval=0.5)
                    sleep(0.5+cooldown)
                    press('enter')
                    sleep(0.5+cooldown)
                    press('esc')
                    sleep(0.5+cooldown)
                    if choice == 'rcs':
                        rcs=True                 
                    elif choice == 'rcg':
                        rcg=True
        else:
            "Кнопку провести и закрыть"
            for _ in range(5):
                press('tab')
                sleep(0.5+cooldown)  
            press('enter')
            sleep(10+cooldown) 
            
    _case_to_court=False
    _case_to_public=False
    "Кнопка передать в суд — зелёная (204,255,216), допуск ±15"
    pos = find_button_center_in_region(BUTTON_PANEL_REGION, GREEN_TRANSFER_RGB, tolerance_abs=15, min_pixels=10)
    if pos:
        _case_to_court=True
        click(pos[0], pos[1])
        sleep(0.5 + cooldown)
        "Кнопка дело в суде нажата"
        sleep(20 + cooldown)
        
    #Долгое ожидание перед второй кнопкой
    sleep(30+cooldown)
    # Кнопка «В пользу общества» (зелёная) или «Не в пользу общества» (красная)
    target_rgb = RED_BUTTON_RGB if result_case == "отк" else GREEN_BUTTON_RGB
    pos = find_button_center_in_region(BUTTON_PANEL_REGION, target_rgb, tolerance_abs=15 if "удп" != "отк" else None, min_pixels=10)
    if pos:
        _case_to_public=True
        click(pos[0], pos[1])
        sleep(0.5 + cooldown)
        "Кнопка с обществом нажата"
        sleep(20 + cooldown)
        
    
    sleep(25+cooldown)
    press('esc')
    sleep(10+cooldown)
    press('enter')
    sleep(15+cooldown)
    
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
   
