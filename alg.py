# Загружаем xlsx с отчётом
from turtle import position
import pyautogui as pg
from utils import load_from_xlsx, addToBuffer, getFromBuffer, selectAll
data = load_from_xlsx("27022026.xlsx")

#Поиск дела этап 1
test_session = "48-090909"
# addToBuffer(test_session)
selectAll()
#Для начала закрываем все вкладки, после чего программа готова к поиску дела
