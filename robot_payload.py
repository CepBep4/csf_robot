"""
Форматы данных для robot.interface_start(mode, data).

Фронт может присылать данные в «сыром» виде (например, строка из Excel с русскими заголовками).
Нормализатор приводит их к виду, ожидаемому в interface_start.
"""

# Маппинг заголовков из документа выгрузки (Excel) в ключи для режима "set"
DOC_ROW_TO_SET_KEYS = {
    "ID дела  (ID/ номер убытка)": "number_case",
    "Наименование должника": "name_defedant",
    "Суд": "court",
    "Дата вынесения решения": "date_base",
    "Дата вступления в силу": "date_plus_mounth",
    "Результат судебного дела": "result_case",
    "Запрошенная сумма (страховое)": "summ_requests_s",
    "Взысканная сумма (основной долг)": "summ_real_s",
    "Сумма госпошлины (запрошено)": "summ_requests_g",
    "Сумма госпошлины": "summ_real_g",
    "Вид исполнительного листа": "view_ip_list",
    "Номер исполнительного листа": "number_ip_list",
    "Сумма долга": "summ",
    "Дата получения исполнительного листа": "data_get_ip_list",
}


def normalize_robot_payload(mode: str, raw: dict | list) -> dict | list:
    """
    Приводит сырые данные с фронта к формату для interface_start(mode, data).

    Режимы и ожидаемый результат:
    - check:  data = {"number_case": str}
    - set:    data = {number_case, name_defedant, court, date_base, date_plus_mounth,
              result_case, summ_requests_s, summ_real_s, summ_requests_g, summ_real_g,
              view_ip_list, number_ip_list, summ, data_get_ip_list, cooldown?}
    - download: data = {"number_case": str}
    - change_filename: data = {"path": str, "mode": str}
    """
    if mode == "check":
        if isinstance(raw, dict):
            return {"number_case": str(raw.get("number_case", raw.get("ID дела  (ID/ номер убытка)", "")))}
        if isinstance(raw, list):
            return [normalize_robot_payload("check", r if isinstance(r, dict) else {"number_case": r}) for r in raw]
        return {"number_case": ""}

    if mode == "set":
        if isinstance(raw, list):
            return [normalize_robot_payload("set", item) for item in raw if isinstance(item, dict)]
        if not isinstance(raw, dict):
            return raw
        out = {}
        for doc_header, key in DOC_ROW_TO_SET_KEYS.items():
            out[key] = raw.get(key, raw.get(doc_header))
        for key in ("number_case", "name_defedant", "court", "date_base", "date_plus_mounth",
                    "result_case", "summ_requests_s", "summ_real_s", "summ_requests_g", "summ_real_g",
                    "view_ip_list", "number_ip_list", "summ", "data_get_ip_list", "cooldown"):
            if key in raw and key not in out:
                out[key] = raw[key]
        return out

    if mode == "download":
        if isinstance(raw, dict):
            return {"number_case": str(raw.get("number_case", raw.get("ID дела  (ID/ номер убытка)", "")))}
        if isinstance(raw, list):
            return [normalize_robot_payload("download", r if isinstance(r, dict) else {"number_case": r}) for r in raw]
        return {"number_case": ""}

    if mode == "change_filename":
        if not isinstance(raw, dict):
            return raw
        return {
            "path": str(raw.get("path", raw.get("fixPath", ""))),
            "mode": str(raw.get("mode", raw.get("fixMode", "from1c"))),
        }

    return raw
