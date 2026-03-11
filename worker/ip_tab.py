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
GREEN_IP_RGB = (156, 250, 156)

def ip_tab(
    view_ip_list: str, #Просто строка с содержанием
    number_ip_list: str, #В формате ФС №XXXX
    summ: float, #Сумма через xxx.yy
    data_get_ip_list: str, #Дата в формате день.месяц.год
    cooldown = 0
):
    _ip_received=False
    ip_list_map = {
        "Оригинал исполнительного листа":3,
        "Постановление ФССП":4,
        "Инкассове списание из банка":2,
        "Дубликат испольнительного листа":1,
        "Добровольная оплата по решению суда":0  
    }

    if view_ip_list not in ip_list_map:
        return ("Информация ВИД ИСПОЛНИТЕЛЬНОГО ЛИСТА заполнена неверно", False)
    
    "Открываем вкладу ИП"
    for _ in range(26):
        press('tab')
        sleep(0.1+cooldown)
        
    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer() == 'none':
        return ("Вкладка ИП не открылась", False)
    
    hotkey('ctrl', 'shift', 'f4', interval=0.5)
    sleep(3+cooldown)
    
    "Переходим к виду исполнительного листа"
    for _ in range(2):
        press("tab")
        sleep(0.1+cooldown)
        
    for _ in range(7):
        press('up')
        sleep(0.1+cooldown)
        
    "Указываем нужный вид"
    for _ in range(ip_list_map[view_ip_list]):
        press('down')
        sleep(0.1+cooldown)
    press('enter')
    sleep(2+cooldown)
    
    "Переходим в исполнительный лист"
    for _ in range(2):
        press("tab")
        sleep(0.1+cooldown)  
    
    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)
    
    if getFromBuffer() == 'none':
        "Вводим номер исл"
        addToBuffer(number_ip_list)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
    "Переходем к дате выдачи исл"
    for _ in range(2):
        press("tab")
        sleep(0.1+cooldown)  
        
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
        addToBuffer(data_get_ip_list)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)
        
    "Переходем к сумме"
    for _ in range(2):
        press("tab")
        sleep(0.1+cooldown) 
        
    "Выделяем всё, чтобы скопировать"
    hotkey('ctrl', 'a', interval=0.5)
    sleep(2+cooldown)

    "Копируем, чтобы проверить по буфферу"
    addToBuffer("none")
    sleep(1+cooldown)
    hotkey('ctrl', 'c', interval=0.5)
    sleep(1+cooldown)

    if getFromBuffer() == '0,00':
        "Вводим сумму"
        addToBuffer(summ)
        sleep(1+cooldown)
        hotkey('ctrl', 'v', interval=0.5)
        sleep(2+cooldown)

    for _  in range(3):
        press('tab')
        sleep(0.1+cooldown)
    press('enter')
    sleep(20+cooldown)
    
    "Кнопка передать в суд — зелёная (204,255,216), допуск ±15"
    pos = find_button_center_in_region(BUTTON_PANEL_REGION, GREEN_IP_RGB, tolerance_abs=15, min_pixels=10)
    if pos:
        _ip_received=True
        click(pos[0], pos[1])
        sleep(0.5 + cooldown)
        "Кнопка ИП получено"
        sleep(20 + cooldown)
        
    for _ in range(3):
        press('esc')
        sleep(10+cooldown)
    
    return (f"Вкладка ИП успешно заполнена. Кнопка ИП получен нажата:{_ip_received}", True)
    
def ip_tab_check(cooldown = 0):
    "Открываем вкладу ИП"
    click(1687,774)
    sleep(2+cooldown)
    
    "Проверяем открылась ли вкладка"
    if pixel(1687,774) == (255,255,255):        
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
        number_ip_list = getFromBuffer()
            
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
        summ = getFromBuffer()
            
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
        data_get_ip_list = getFromBuffer()
            
        "Закрываем вкладку"
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)
            
        return (f"Вкладка ИП успешно проверена", True, {
            'number_ip_list':number_ip_list,
            'summ':summ,
            'data_get_ip_list':data_get_ip_list
        })
    else:
        "Закрываем вкладку"
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)   
        return ("Вкладка ИП не открылась", False)