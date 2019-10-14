from collections import Counter
import re
import xlwt

allwords = list()                                       # Объявление списка, содержащего все слова
print("Введите путь до файла:")
path = str(input())                                     # Получение пути до файла с текстом через консоль
fileData = open(path)                                   # Считывание всех строк из файла
for line in fileData:                                   # Цикл для проверки каждой строки
    newStroka = line.replace("\n","")                   # Удаление лишнего отступа
    newStroka = re.sub('[\\\\/*?".<>|]', '', newStroka)  # Регулярное выражение удаляющее символы
    newStroka = re.sub('[\\\\:;]', ' ', newStroka)      # Регулярное выражение заменяющее ; и : на пробел
    stroka = re.split(r'\s+|,\s*', newStroka)           # Разделение строки на отдельные слова
    for word in stroka:                                 # Цикл добавляющий все слова в строке в общий список слов
        allwords.append(word.lower())

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')

cellsCounter = 0
counter = Counter(allwords)                             # Словарь, подсчитывающий кол-во повторяющихся значений
for pair in counter.most_common():
    ws.write(cellsCounter, 0, pair[0])
    ws.write(cellsCounter, 1, pair[1])
    cellsCounter = cellsCounter + 1

wb.save('Words.xls')



