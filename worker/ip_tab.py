from pyautogui import click, press, pixel, hotkey
from time import sleep
from worker.utils import addToBuffer, getFromBuffer

def ip_tab(
    view_ip_list: str, #Просто строка с содержанием
    number_ip_list: str, #В формате ФС №XXXX
    summ: float, #Сумма через xxx.yy
    data_get_ip_list: str, #Дата в формате день.месяц.год
    cooldown = 0
):
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
        
        _transfer_to_FSSP = False
        "Передаём в ФССП"
        if pixel(2159,1344) == (145, 239, 145) and pixel(1952,1344) == (135, 206, 250):
            _transfer_to_FSSP = True
            "Нажимаем на кнопку"
            click(2159,1344)
            sleep(25+cooldown)
            
        "Закрываем вкладку"
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)
            
        return (f"Вкладка ИП успешно заполнена. Кнопка Передать в ФССП прожата:{_transfer_to_FSSP}", True)
    else:
        "Закрываем вкладку"
        press('esc') 
        sleep(2+cooldown)   
        press('esc') 
        sleep(2+cooldown)   
        return ("Вкладка ИП не открылась", False)
    
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