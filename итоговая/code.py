import csv
import os
import json

from openpyxl import Workbook
from prettytable import PrettyTable
import matplotlib.pyplot as plt    

def showGraph():
    table1 = dataFromCsv('table1.txt')
    table2 = dataFromCsv('table2.txt')
    priceList = []
    for i in table2[1:]:
        priceList.append(int(i[2]))
    
    priceList.append(0)
    priceList.append(0)
    
    priceList.sort()
                    
    quantityList = []
    for i in table1[1:]:
        quantityList.append(int(i[3]))

    quantityList.sort()

    plt.plot(priceList, quantityList)
    plt.xlabel("Ціна")
    plt.ylabel("Кількість")
    plt.show()


def dataFromCsv(csvFilePath):
    data = []
    with open(csvFilePath, encoding='utf-8') as csvf:
        csvReader = csv.reader(csvf, delimiter=',')
        for row in csvReader:
            data.append(row)

    return data

def writeExcelFromPT(pt, filePath):
    wb = Workbook()
    
    ws = wb.active

    data = pt.rows

    for row in data:
        ws.append(row)

    wb.save(filename = filePath)

def writePrettyTable(pt, filePath):
    with open(filePath, 'wt', encoding='utf-8') as fout:
        fout.write(pt.get_string())

def setUpTable():
    data1 = dataFromCsv('table1.txt')
    data2 = dataFromCsv('table2.txt')
    pt = PrettyTable(title="Замовлення на продаж методичної літератури та програмних засобів по магазину")

    # 1 column
    columns = []

    for i in data2[1:]:
        columns.append(i[1])

    columns.append(' ')
    columns.append(' ')
    pt.add_column("Найменування товара", columns)

    # 2 column
    columns = []

    for i in data1[1:]:
        columns.append(i[1])

    pt.add_column("Номер заказа", columns)
    
    # 3 column
    columns = []

    for i in data1[1:]:
        columns.append(i[0])

    pt.add_column("Код клієнта", columns)

    # 4 column
    columns = []

    for i in data1[1:]:
        columns.append(i[3])

    pt.add_column("Кількість", columns)

    # 5 column
    columns = []

    for i in data2[1:]:
        columns.append(i[2])
    columns.append(' ')
    columns.append(' ')

    pt.add_column("Ціна", columns)

    # 6 column
    columns = []

    for i, j in list(zip(data1[1:], data2[1:])):
        columns.append(str(int(i[3])*int(j[2])))
    columns.append(' ')
    columns.append(' ')

    pt.add_column("Сума", columns)
    
    return pt

def writeJsonFromPT(pt, jsonFilePath):
    with open(jsonFilePath, 'wt', encoding='utf-8') as fout:
        fout.write(pt.get_json_string())

def chooseCategory(pt, name):
    try:
        name = input("Критерій: ").capitalize()
    except AttributeNotFound:
        return
    else:
        print(pt.get_string(fields=[name]))
    

def main():
    pt = setUpTable()
    while True:
        print(pt)
        print("""
1.створити текстовий файл
2.створити файл json
3.створити файл ексель.
4. Вивести графік на екран
5. Вибрати значення за критерієм
""")
        try:
            a = int(input("Вибрати дію (цифра): "))
        except ValueError:
            input("Неправильний ввід, спробуйте ще раз.")
            continue
        else:
            match a:
                case 1:
                    writePrettyTable(pt, "таблиця.txt")
                case 2:
                    writeJsonFromPT(pt, "таблиця.json")
                case 3:
                    writeExcelFromPT(pt, "таблиця.xlsx")
                case 4:
                    showGraph()
                case 5:
                    
                    chooseCategory(pt, "Ціна")

        input("\nГотово.\n")



if __name__ == '__main__':
    main()
    
