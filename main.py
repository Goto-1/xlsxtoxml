import openpyxl
import xml.etree.ElementTree as ET

# Define variable to load the wookbook
wookbook = openpyxl.load_workbook("Контактная карта АО ПремьерСтрой_2023.05.16.xlsx")
#d = {'Бованенковское месторождение':'БВН','Восточно-Мессояхское месторождение':'МЕС','Находкинское месторождение':'Нхд','Опорная база г. Коротчаево':'Кор','Опорная база с. Газ-Сале':'Газ','Офис г. Москва':'МСК','Офис г. Тюмень':'ТМН','Тазовское месторождение':'ТАЗ','Харасавэйское месторождение':'ХРСВ','Центр Аудио-Конференций':'АКС'}
print(wookbook.sheetnames)