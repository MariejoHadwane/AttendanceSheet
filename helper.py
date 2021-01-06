import xlrd
import datetime
import random
from datetime import date, timedelta, datetime
import pandas as pd
import csv

def getEmployeesData(csvPath):
    employeesData = {}
    with open(csvPath) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count != 0:
                employeesData[row[0]] = int(row[1])
            line_count += 1
    return employeesData

def daysCurrentMonth():
    m = datetime.now().month
    y = datetime.now().year
    ndays = (date(y, m+1, 1) - date(y, m, 1)).days
    d1 = date(y, m, 1)
    d2 = date(y, m, ndays)
    delta = d2 - d1

    fullDates=[(d1 + timedelta(days=i)).strftime('%d-%m-%Y') for i in range(delta.days + 1)]
    return fullDates

def getEmployeesToCome(employeesEvenOdd, maxPercentage, fullDate, oddEvenDate):
    result = {}
    for j in fullDate:
        list = []
        dayOfWeek=datetime.strptime(j,'%d-%m-%Y').strftime('%A')
        if(dayOfWeek!= "Saturday" and dayOfWeek!="Sunday"):
            evenOddValue = oddEvenDate[dayOfWeek]
            for n in employeesEvenOdd:
                if (evenOddValue == employeesEvenOdd[n]):
                    list.append(n)
            random.shuffle(list)
            maxNumber = maxPercentage * len(employeesEvenOdd)

            result[j] = list[: int(maxNumber)]
    return result

def str_len(str):
    try:
        row_l=len(str)
        utf8_l=len(str.encode('utf-8'))
        return (utf8_l-row_l)/2+row_l
    except:
        return None
    return None

def createExcelFromDictionary(dictionary, excelPath):
    writer = pd.ExcelWriter(excelPath, engine = 'xlsxwriter')
    df = pd.DataFrame.from_dict(dictionary, orient = 'index').transpose()
    df.to_excel(writer, sheet_name = "Sheet1")
    worksheet = writer.sheets["Sheet1"]
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((
            series.astype(str).map(str_len).max(),
            str_len(str(series.name))
        )) + 2
        worksheet.set_column(idx, idx, max_len)
    writer.save()
