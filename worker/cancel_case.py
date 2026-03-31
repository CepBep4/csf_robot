from pyautogui import click, press, pixel
from time import sleep
from worker import BUTTON_PANEL_REGION

from pyautogui import click, press, pixel, screenshot
from time import sleep

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

# Зелёная «Передать в суд» (другой оттенок)
CANCEL_CASE = (253, 162, 121)

def cansel_case(cooldown=0):
    pos = find_button_center_in_region(BUTTON_PANEL_REGION, CANCEL_CASE, tolerance_abs=15, min_pixels=10)
    if pos:
        click(pos[0], pos[1])
        "Кнопка ИП получено"
        sleep(30+cooldown)
        press('esc')
        sleep(10+cooldown)
        return ("Дело успешно отменено", True)
    else:
        press('esc')
        sleep(10+cooldown)
        press('esc')
        sleep(10+cooldown)
        return ("Кнопка не обнаружена", False)

# print(pixel(2399,1342))
# (253, 162, 121)

# 000037838
# 000037977
# 000038373
# 000038375
# 000038380
# 000038521
# 000038583
# 000038677
# 000038792

