from time import sleep

from pyautogui import click, press, screenshot
from worker import BUTTON_PANEL_REGION


def _pixel_matches(img, px, py, target_rgb, tol):
    if px < 0 or py < 0 or px >= img.size[0] or py >= img.size[1]:
        return False
    r, g, b = img.getpixel((px, py))[:3]
    tr, tg, tb = target_rgb
    return abs(r - tr) <= tol and abs(g - tg) <= tol and abs(b - tb) <= tol


def find_button_center_in_region(region, target_rgb, tolerance_abs=15, min_pixels=10):
    """
    Ищет в области первое связное пятно кнопки target_rgb и возвращает его центр.
    region: (left, top, width, height)
    """
    img = screenshot(region=region)
    w, h = img.size
    left, top = region[0], region[1]
    tol = int(tolerance_abs)

    def matches(px, py):
        return _pixel_matches(img, px, py, target_rgb, tol)

    for py in range(h):
        for px in range(w):
            if not matches(px, py):
                continue
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


PRETRIAL_BUTTON_RGB = (232, 232, 232)


def set_pretrial(cooldown=0):
    """
    Прожимает кнопку "выставлена досудебка" по цвету из alg.py.
    """
    pos = find_button_center_in_region(
        BUTTON_PANEL_REGION,
        PRETRIAL_BUTTON_RGB,
        tolerance_abs=15,
        min_pixels=10,
    )
    if not pos:
        return ("Кнопка 'выставлена досудебка' не обнаружена", False)
    click(pos[0], pos[1])
    sleep(30 + cooldown)
    press("esc")
    sleep(10 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    press("esc")
    sleep(5 + cooldown)
    return ("Шаг 'выставлена досудебка' выполнен", True)
