from worker.search_case import search_case
from worker.court_tab import court_tab
from worker.ip_tab import ip_tab
from worker.cancel_case import cansel_case
from worker.pretrial_tab import set_pretrial as set_pretrial_action
from worker.case_to_court import press_case_to_court as press_case_to_court_action
from worker.in_favor_society import press_in_favor_society as press_in_favor_society_action
from worker.sent_to_fssp import press_sent_to_fssp as press_sent_to_fssp_action
from time import sleep


def _required(case: dict, key: str):
    value = case.get(key)
    if value is None:
        raise ValueError(f"Не передано поле '{key}'")
    return value


def set_pretrial(case: dict, cooldown: float = 0):
    return set_pretrial_action(cooldown=cooldown)


def open_case(case: dict, cooldown: float = 0):
    number_case = str(_required(case, "number_case")).strip()
    return search_case(number_case, cooldown=cooldown)


def fill_court_tab(case: dict, cooldown: float = 0):
    return court_tab(
        name_defedant=_required(case, "name_defedant"),
        court=_required(case, "court"),
        date_base=_required(case, "date_base"),
        date_plus_mounth=_required(case, "date_plus_mounth"),
        result_case=_required(case, "result_case"),
        summ_requests_s=_required(case, "summ_requests_s"),
        summ_real_s=_required(case, "summ_real_s"),
        summ_requests_g=_required(case, "summ_requests_g"),
        summ_real_g=_required(case, "summ_real_g"),
        cooldown=cooldown,
        one_tab=True,
    )


def press_in_favor_society(case: dict, cooldown: float = 0):
    result_case = _required(case, "result_case")
    return press_in_favor_society_action(result_case=result_case, cooldown=cooldown)


def press_case_to_court(case: dict, cooldown: float = 0):
    return press_case_to_court_action(cooldown=cooldown)


def fill_ip_tab(case: dict, cooldown: float = 0):
    return ip_tab(
        view_ip_list=_required(case, "view_ip_list"),
        number_ip_list=_required(case, "number_ip_list"),
        summ=float(_required(case, "summ")) + float(_required(case, "summ_real_g")),
        data_get_ip_list=_required(case, "data_get_ip_list"),
        cooldown=cooldown,
    )


def press_sent_to_fssp(case: dict, cooldown: float = 0):
    return press_sent_to_fssp_action(cooldown=cooldown)


def close_case(case: dict, cooldown: float = 0):
    return cansel_case(cooldown=cooldown)


MODE_MAP = {
    "Прожать выставлена досудебка": set_pretrial,
    "Открыть дело": open_case,
    "Заполнить вкладку суд": fill_court_tab,
    "Прожать в пользу общества": press_in_favor_society,
    "Прожать дело в суд": press_case_to_court,
    "Заполнить ИП": fill_ip_tab,
    "Прожать передано в ФССП": press_sent_to_fssp,
    "Прожать закрыть дело": close_case,
}

MODE_PROD_MAP = {
    # 06_Ожидание итогов от аутсорсеров
    "06_Ожидание итогов от аутсорсеров": [
        "Открыть дело",
        "Прожать выставлена досудебка"
    ],
    # 07_Ожидание ответа по досудебному требованию
    "07_Ожидание ответа по досудебному требованию": [
        "Открыть дело",
        "Прожать дело в суд",
    ],
    # 09_Ожидание решения суда
    "09_Ожидание решения суда": [
        "Открыть дело",
        "Заполнить вкладку суд",
        "Прожать в пользу общества",
    ],
    # 11_Получение ИЛ
    "11_Получение ИЛ": [
        "Открыть дело",
        "Заполнить ИП",
        "Прожать передано в ФССП",
    ],
    # 13_Исполнение ИЛ в ФССП
    "13_Исполнение ИЛ в ФССП": [
        "Открыть дело",
        "Прожать закрыть дело",
    ],
}


def run_prod_mode(case: dict, mode_key: str, cooldown: float = 0) -> list[dict]:
    """
    Выполняет шаги по ключу прод-режима из MODE_PROD_MAP.
    Возвращает список шагов: [{"step": str, "ok": bool, "message": str}, ...]
    """
    steps = MODE_PROD_MAP.get(mode_key) or []
    results = []
    if not steps:
        return [{"step": "", "ok": False, "message": f"Неизвестный режим: {mode_key}"}]

    stage_delay_sec = 0.0
    try:
        stage_delay_sec = float(case.get("stageDelaySec", 0) or 0)
    except (TypeError, ValueError):
        stage_delay_sec = 0.0
    if stage_delay_sec < 0:
        stage_delay_sec = 0.0

    for idx, step_name in enumerate(steps):
        fn = MODE_MAP.get(step_name)
        if fn is None:
            results.append({"step": step_name, "ok": False, "message": "Шаг не реализован"})
            break
        try:
            msg, ok = fn(case, cooldown=cooldown)
        except Exception as e:
            results.append({"step": step_name, "ok": False, "message": str(e)})
            break
        results.append({"step": step_name, "ok": bool(ok), "message": str(msg)})
        if not ok:
            break
        if idx < len(steps) - 1 and stage_delay_sec > 0:
            sleep(stage_delay_sec)
    return results

