import calendar
from datetime import datetime
from robot_log import get_logger
from robot_state import is_stop_requested
from worker import (
    hande_case_by_setting_information_to_1c,
    hande_case_by_downloading_information_from_1c,
    handle_cancel_case_into_1c,
)


def validate_before_setting(case: dict) -> dict:
    #Формируем дату + месяц
    date_base = case.get('date_base')
    if date_base is None:
        raise ValueError('Поле «Дата вынесения решения» не заполнено')
    date_base = date_base.strip()
    try:
        dt = datetime.strptime(date_base, "%d.%m.%Y")
    except ValueError as e:
        msg = str(e)
        if "day is out of range for month" in msg:
            msg = "день вне диапазона для месяца (например 31.04 или 29.02 в невисокосном году)"
        elif "time data" in msg and "does not match format" in msg:
            msg = "неверный формат даты (ожидается ДД.ММ.ГГГГ)"
        raise ValueError(f'Поле «Дата вынесения решения»: значение «{date_base}» — {msg}') from e
    # +1 месяц: если день не существует в следующем месяце — берём последний день месяца
    # (например 31.10 → 30.11, 30.01 → 28.02 или 29.02)
    next_month = dt.month % 12 + 1
    next_year = dt.year + (1 if dt.month == 12 else 0)
    last_day = calendar.monthrange(next_year, next_month)[1]
    day = min(dt.day, last_day)
    dt = dt.replace(year=next_year, month=next_month, day=day)
    case["date_plus_mounth"] = dt.strftime("%d.%m.%Y")
    
    #Формируем result_case
    result_case = case.get('result_case')
    if result_case is None:
        raise ValueError('Поле «Результат судебного дела» не заполнено')
    result_case = str(result_case).strip().lower()
    if result_case == "удовлетворены":
        case['result_case'] = 'удп'
    elif "удовлетворены" in result_case:
        case['result_case'] = 'удч'
    else:
        case['result_case'] = 'отк'
        
    #Правим summ_real_s
    val_s = case.get('summ_real_s')
    if val_s is None or (isinstance(val_s, str) and not val_s.strip()):
        raise ValueError('Поле «Взысканная сумма (основной долг)» не заполнено')
    try:
        case['summ_real_s'] = float(''.join(str(val_s).replace(',', '.').split()))
    except (ValueError, TypeError) as e:
        raise ValueError(f'Поле «Взысканная сумма (основной долг)»: значение «{val_s}» — не число') from e
    case['summ_requests_s'] = case['summ_real_s']

    #Правим summ_real_g
    val_g = case.get('summ_real_g')
    if val_g is None or (isinstance(val_g, str) and not val_g.strip()):
        raise ValueError('Поле «Сумма госпошлины» не заполнено')
    try:
        case['summ_real_g'] = float(''.join(str(val_g).replace(',', '.').split()))
    except (ValueError, TypeError) as e:
        raise ValueError(f'Поле «Сумма госпошлины»: значение «{val_g}» — не число') from e
    case['summ_requests_g'] = case['summ_real_g']

    #Заполняем поле view_ip_list
    case['view_ip_list'] = "Оригинал исполнительного листа"

    #Приводим к нужному формату поле number_ip_list
    num_ip = case.get('number_ip_list')
    if num_ip is None or (isinstance(num_ip, str) and not num_ip.strip()):
        raise ValueError('Поле «Номер исполнительного листа» не заполнено')
    num_ip = str(num_ip).strip()
    if len(num_ip) < 5:
        raise ValueError(f'Поле «Номер исполнительного листа»: значение «{num_ip}» — слишком короткое')
    case['number_ip_list'] = f'ФС №{"".join(num_ip[5::])}'

    #Приводим поле summ
    val_sum = case.get('summ')
    if val_sum is None or (isinstance(val_sum, str) and not val_sum.strip()):
        raise ValueError('Поле «Сумма долга» не заполнено')
    try:
        case['summ'] = float(''.join(str(val_sum).replace(',', '.').split()))
    except (ValueError, TypeError) as e:
        raise ValueError(f'Поле «Сумма долга»: значение «{val_sum}» — не число') from e
    
    return case

def interface_start(mode: str, data: dict | list):
    log = get_logger()
    log.info("Старт: mode=%s", mode)

    try:
        if mode == "check":
            pass

        elif mode == "set":
            from worker.modeAll import MODE_PROD_MAP, run_prod_mode
            if not isinstance(data, list):
                data = [data]
            results = []
            for i, case in enumerate(data):
                if is_stop_requested():
                    log.warning("Остановка по запросу пользователя после %s из %s дел", i, len(data))
                    break
                validate_before_run = bool(case.get("validateBeforeRun", True))
                if validate_before_run:
                    try:
                        case = validate_before_setting(case)
                    except ValueError as e:
                        log.error("Дело %s: ошибка валидации — %s", case.get("number_case", i + 1), e)
                        results.append({"number_case": case.get("number_case"), "ok": False, "message": str(e)})
                        continue
                set_info = case.get("setInfo") if isinstance(case, dict) else None
                set_info = str(set_info).strip().lower() if set_info is not None else ""
                # Приводим ключ к регистру словаря прод-режимов
                prod_mode_map_ci = {str(k).strip().lower(): k for k in MODE_PROD_MAP.keys()}
                if set_info not in ("court", "ip", "both") and set_info not in prod_mode_map_ci:
                    msg = "Не выбран режим заполнения (setInfo)."
                    log.error("Дело %s: %s", case.get("number_case", i + 1), msg)
                    results.append({"number_case": case.get("number_case"), "ok": False, "message": msg})
                    continue
                log.info("Обработка дела %s (%s из %s)", case.get("number_case"), i + 1, len(data))
                if set_info in prod_mode_map_ci:
                    mode_key = prod_mode_map_ci[set_info]
                    step_logs = run_prod_mode(case, mode_key=mode_key)
                    ok = len(step_logs) > 0 and all(bool(s.get("ok")) for s in step_logs)
                    if step_logs:
                        msg = step_logs[-1].get("message", "")
                    else:
                        msg = "Пустой план режима"
                    results.append({
                        "number_case": case.get("number_case"),
                        "ok": ok,
                        "message": msg,
                        "mode": mode_key,
                        "steps": step_logs,
                    })
                else:
                    msg = hande_case_by_setting_information_to_1c(case, set_info)
                    results.append({"number_case": case.get("number_case"), "ok": "успешно" in msg or "завершён" in msg.lower(), "message": msg})
                log.info("Дело %s: %s", case.get("number_case"), msg)
            return results
        elif mode == "download":
            if not isinstance(data, list):
                data = [data]
            results = []
            for i, item in enumerate(data):
                if is_stop_requested():
                    log.warning("Остановка по запросу пользователя после %s из %s дел", i, len(data))
                    break
                number_case = item.get("number_case", item) if isinstance(item, dict) else item
                number_case = str(number_case).strip()
                if not number_case:
                    log.warning("Пропуск дела с пустым номером (индекс %s)", i + 1)
                    results.append({"number_case": "", "ok": False, "message": "Пустой номер дела"})
                    continue
                log.info("Скачивание дела %s (%s из %s)", number_case, i + 1, len(data))
                try:
                    hande_case_by_downloading_information_from_1c(number_case)
                    results.append({"number_case": number_case, "ok": True, "message": "Скачано"})
                except Exception as e:
                    log.exception("Дело %s: ошибка при скачивании — %s", number_case, e)
                    results.append({"number_case": number_case, "ok": False, "message": str(e)})
            return results

        elif mode == "cancel_case":
            if not isinstance(data, list):
                data = [data]
            results = []
            for i, item in enumerate(data):
                if is_stop_requested():
                    log.warning("Остановка по запросу пользователя после %s из %s дел", i, len(data))
                    break
                number_case = item.get("number_case", item) if isinstance(item, dict) else item
                number_case = str(number_case).strip()
                if not number_case:
                    log.warning("Пропуск дела с пустым номером (индекс %s)", i + 1)
                    results.append({"number_case": "", "ok": False, "message": "Пустой номер дела"})
                    continue
                log.info("Отмена дела %s (%s из %s)", number_case, i + 1, len(data))
                try:
                    handle_cancel_case_into_1c(number_case)
                    results.append({"number_case": number_case, "ok": True, "message": "Отменено"})
                except Exception as e:
                    log.exception("Дело %s: ошибка при отмене — %s", number_case, e)
                    results.append({"number_case": number_case, "ok": False, "message": str(e)})
            return results

        log.info("Завершено: mode=%s", mode)
    except Exception as e:
        log.exception("Ошибка при выполнении: %s", e)
        raise
