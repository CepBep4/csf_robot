import json
from pathlib import Path

from robot_state import is_stop_requested
from worker.court_tab import court_tab, court_tab_check
from worker.ip_tab import ip_tab, ip_tab_check
from worker.search_case import search_case
from worker.download_mode import download_mode

_config_path = Path(__file__).resolve().parent.parent / "config.json"
with open(_config_path, encoding="utf-8") as f:
    _config = json.load(f)

SEARCH_COOLDOWN = _config.get("delaySearchCase", 0.0)
COURT_TAB_DELAY = _config.get("delayCourtTab", 0.0)
DOWNLOAD_CASES_DELAY = _config.get("delayDownloadCases", 180.0)
STAGE_SWITCH_DELAY = _config.get("delayStageSwitch", 10.0)

"Запуск заполнения 1С"
def hande_case_by_setting_information_to_1c(data: dict):
    if is_stop_requested():
        return "Остановлено по запросу пользователя"
    #Данные для поиска дела
    number_case: int = data.get("number_case")
        
    #Данные для суда
    name_defedant: str = data.get("name_defedant")
    court: str = data.get("court")
    date_base: str = data.get("date_base")
    date_plus_mounth: str = data.get("date_plus_mounth")
    result_case: str = data.get("result_case")
    summ_requests_s: float = data.get("summ_requests_s")
    summ_real_s: float = data.get("summ_real_s")
    summ_requests_g: float = data.get("summ_requests_g")
    summ_real_g: float = data.get("summ_real_g")
    
    #Данные для ИП
    view_ip_list: str = data.get("view_ip_list")
    number_ip_list: str = data.get("number_ip_list")
    summ: float = data.get("summ")
    data_get_ip_list: str = data.get("data_get_ip_list")
    
    #Ищем дело
    sc = search_case(number_case, cooldown=SEARCH_COOLDOWN)
    print(f"Дело: {number_case}\nИнформация: {sc[0]}")
    if is_stop_requested():
        return "Остановлено по запросу пользователя"
    if sc[1]:
        #Заполняем вкладку суд
        ct = court_tab(
            name_defedant=name_defedant,
            court=court,
            date_base=date_base,
            date_plus_mounth=date_plus_mounth,
            result_case=result_case,
            summ_requests_s=summ_requests_s,
            summ_real_s=summ_real_s,
            summ_requests_g=summ_requests_g,
            summ_real_g=summ_real_g,
            cooldown=COURT_TAB_DELAY,
        )
        print(f"Дело: {number_case}\nИнформация: {ct[0]}")
        if is_stop_requested():
            return "Остановлено по запросу пользователя"
        if ct[1]:
            #Заполняем вкладку исполнительное производство
            it = ip_tab(
                view_ip_list=view_ip_list,
                number_ip_list=number_ip_list,
                summ=summ,
                data_get_ip_list=data_get_ip_list,
                cooldown=COURT_TAB_DELAY,
            )
            print(f"Дело: {number_case}\nИнформация: {it[0]}")
            if it[1]:
                return f"Полный цикл заполнения успешно завершён. {it[0]}"
            else:
                "Обрабатываем ошибки по вкладке ИП"
                return it[0]
        else:
            "Обрабатываем ошибки по вкладке СУД"
            return ct[0]
    else:
        "Обрабатываем ошибки по поиску ДЕЛА"
        return sc[0]
    
    

"Запуск проверки информации в 1С"
def hande_case_by_checking_information_from_1c(number_case: str) -> dict:
    if is_stop_requested():
        return {}
    data_recieve = {}
    
    #Ищем дело
    sc = search_case(number_case, cooldown=SEARCH_COOLDOWN)
    print(f"Дело: {number_case}\nИнформация: {sc[0]}")
    if is_stop_requested():
        return {}
    if sc[1]:
        data_recieve["number_case"] = number_case
        #Заполняем вкладку суд
        ct = court_tab_check(COURT_TAB_DELAY)
        print(f"Дело: {number_case}\nИнформация: {ct[0]}")
        if is_stop_requested():
            return data_recieve
        if ct[1]:
            data_recieve.update(ct[2])
            #Заполняем вкладку исполнительное производство
            it = ip_tab_check(COURT_TAB_DELAY)
            print(f"Дело: {number_case}\nИнформация: {it[0]}")
            if it[1]:
                data_recieve.update(it[2])
                return data_recieve
            else:
                "Обрабатываем ошибки по вкладке ИП"
                return data_recieve
        else:
            "Обрабатываем ошибки по вкладке СУД"
            return ct[0]
    else:
        "Обрабатываем ошибки по поиску ДЕЛА"
        return sc[0]

"Запуск скачивания файлов из 1С"
def hande_case_by_downloading_information_from_1c(number_case: str):
    if is_stop_requested():
        return
    if search_case(number_case)[1]:
        if is_stop_requested():
            return
        download_mode(number_case)