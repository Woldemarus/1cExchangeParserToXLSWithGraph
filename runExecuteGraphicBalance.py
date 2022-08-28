from os import listdir, remove
from os.path import isfile, join
from datetime import datetime
from openpyxl import Workbook, utils
from openpyxl.chart import Reference, LineChart, Series
from openpyxl.chart.axis import DateAxis
import sys
import time

paths = {
        'C:\Project\EdvardPytonSciptLoad\АлфаКом'        
    }

def quarter(date):
    a, b = divmod(date.month, 3)
    return '{}-Q{}'.format(date.year, a + bool(b))
def month(date):
    a, b = divmod(date.month, 1)
    return '{}-{}'.format(date.year, a)


class middleByVypiska(object):
    """docstring"""
    dataPrihod = {}
    dataRashod= {}
    accDate = ""
    accDateR = ""
    counterMonth = 0
    counterMonthR = 0

    def __init__(self):
        """Constructor"""
        self.dataPrihod.clear()
        self.dataRashod.clear()
        self.accDate=""
        self.accDateR=""
        self.counterMonth=0
        self.counterMonthR=0

        self.dataPrihod = {}
        self.dataRashod= {}

        self.accDate = None; del self.accDate
        self.accDateR = None; del self.accDateR
        self.counterMonth = None; del self.counterMonth
        self.counterMonthR = None; del self.counterMonthR
        

    def addValuePrihod(self, value, date):
        """
        рассчитываем среднее на месяц
        """        
        print("Вход")
        print(month(date))
        print(value)
        print(len(self.dataPrihod))
        print(self.accDate)
        print(self.counterMonth)
        

        if self.accDate == "":
            print("Первая запись")
            self.accDate = month(date)            
            self.counterMonth += 1
            self.dataPrihod[self.accDate] = float(value)
        else:
            if self.accDate != month(date):  
                print("Изменение месяца")  
                self.dataPrihod[self.accDate] = self.dataPrihod[self.accDate]/self.counterMonth
                self.dataPrihod[month(date)] = float(value)
                self.counterMonth = 1
                self.accDate = month(date)  
            else:
                if self.accDate == month(date):
                    print("Идет накопление")           
                    self.counterMonth += 1
                    self.dataPrihod[self.accDate] = float(self.dataPrihod[self.accDate]) + float(value)
        print("Результат acc")    
        print(self.dataPrihod[self.accDate])
        print("Результат new")    
        print(self.dataPrihod[month(date)])
    
    def addValueRashod(self, value, date):
        """
        рассчитываем среднее на месяц
        """        
        print("Вход")
        print(month(date))
        print(value)
        print(len(self.dataRashod))
        print(self.accDateR)
        print(self.counterMonthR)
        

        if self.accDateR == "":
            print("Первая запись")
            self.accDateR = month(date)            
            self.counterMonthR += 1
            self.dataRashod[self.accDateR] = float(value)
        else:
            if self.accDateR != month(date):  
                print("Изменение месяца")  
                self.dataRashod[self.accDateR] = self.dataRashod[self.accDateR]/self.counterMonthR
                self.dataRashod[month(date)] = float(value)
                self.counterMonthR = 1
                self.accDateR = month(date)  
            else:
                if self.accDateR == month(date):
                    print("Идет накопление")           
                    self.counterMonthR += 1
                    self.dataRashod[self.accDateR] = float(self.dataRashod[self.accDateR]) + float(value)
        print("Результат acc")    
        print(self.dataRashod[self.accDateR])
        print("Результат new")    
        print(self.dataRashod[month(date)])

    def getTable(self):
        """
        Забираем массив
        """
        self.dataPrihod[list(self.dataPrihod)[-1]] = self.dataPrihod[list(self.dataPrihod)[-1]]/(self.counterMonth)
        self.dataRashod[list(self.dataRashod)[-1]] = self.dataRashod[list(self.dataRashod)[-1]]/(self.counterMonthR)
        self.counterMonth = 0
        self.counterMonthR = 0
        self.accDate = ""
        self.accDateR = ""

        dataBalanse = {}
        for keyP in self.dataPrihod:
            for keyR in self.dataRashod:
                if keyP==keyR:
                    dataBalanse[keyR]=(self.dataPrihod[keyP]) - (self.dataRashod[keyR])
        return self.dataPrihod, self.dataRashod, dataBalanse

    def __del__(self):
        self.dataPrihod.clear()
        self.dataRashod.clear()
        self.accDate=""
        self.accDateR=""
        self.counterMonth=0
        self.counterMonthR=0
        self.accDate = None; del self.accDate
        self.accDateR = None; del self.accDateR
        self.counterMonth = None; del self.counterMonth
        self.counterMonthR = None; del self.counterMonthR

def clear_old_report(path):
    fileList = listdir( path )
    for item in fileList:
        if item.endswith(".xlsx") and item.startswith("report_"):
            remove( join( path, item ) )


def collect_pay_info(file):
    translators = {
        'Дата': 'docDate',
        'ДатаСписано': 'discardedDate',
        'ДатаПоступило': 'receivedDate',        
        'ДатаНачала': 'beginDate',
        'ДатаКонца': 'endDate',
        'НачальныйОстаток': 'beginAmount',
        'ВсегоПоступило': 'totalReceived',
        'ВсегоСписано': 'totalDiscarded',
        'КонечныйОстаток': 'endAmount',
        'Сумма': 'paySum',
        'СекцияДокумент': 'docType',
    }
    metapack = []
    docpack = []
    curdoc = {}
    metadoc = {}
    i=0

    text = open(file, 'r', encoding='Windows-1251').read()
   
    for line in text.splitlines()[1:]:
        i+=1
        if i<19:
            if line.startswith('1CClientBankExchange') or line.startswith('СекцияРасчСчет'):
                continue
            if line.startswith('КонецРасчСчет'):
                metapack.append(metadoc)
                metadoc ={}
                continue
            key, value = line.split('=', maxsplit=1)
            #value = formaters.get(key, value)
            key = translators.get(key, None)
            if key:
                metadoc[key] = value
            continue
        if line.startswith('КонецДокумента'):
            docpack.append(curdoc)
            curdoc = {}
            continue
        if line.startswith('КонецФайла'):
            break
        key, value = line.split('=', maxsplit=1)

        #valueFormated = formaters.get(key, value)
        keyTranslated = translators.get(key, None)
        if keyTranslated:
            curdoc[keyTranslated] = value
    return docpack, metapack




fileIterator = 0
fileIterator4excel = 0
cursorDataPacket = 0


for line in paths:
    #чистим старые репорты
    clear_old_report(line)

    dataPacket ={}
    metaPacket ={}
    onlyfiles = [f for f in listdir(line) if isfile(join(line, f))]   
    #загружаем файлы
    for file4parsing in onlyfiles:
        dataPacket[fileIterator] = collect_pay_info(line + "\\" + file4parsing)
        fileIterator+=1
    #print(dataPacket)
    #print(dataPacket)
    #создаем эксельку
    filename = 'report_' + datetime.now().strftime('%m-%d-%y_%H-%M-%S-%f') + '.xlsx'
    workbook = Workbook()
    workbook.guess_types = True
    sheet = workbook.active


    sheet["D1"] = "Файл"
    sheet["E1"] = "Тип документа"
    sheet["F1"] = "Дата платежа"
    sheet["G1"] = "Приход"
    sheet["H1"] = "Расход"
    sheet["I1"] = "Квартал"
    
    #print(dataPacket)
    cursorCell = 0
    prihod = 0
    rashod = 0
    countFiles = len(dataPacket)
    middlePacketByStatement = []
    middleVipiska = middleByVypiska()
    #print(countFiles)
    for tpm in range(countFiles):
        #print("tpm:" + str(tpm))
        #print("len(dataPacket):" + str(len(dataPacket)))
        #print(dataPacket)
        #print(dataPacket[tpm + cursorDataPacket][1][0]['beginAmount'])
        sheet["A"+str(2+cursorCell)] = "ДатаОт"
        sheet["B"+str(2+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['beginDate']
        sheet["A"+str(3+cursorCell)] = "ДатаДо"
        sheet["B"+str(3+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['endDate']
        sheet["A"+str(4+cursorCell)] = "НачОстаток"
        sheet["B"+str(4+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['beginAmount']
        sheet["A"+str(5+cursorCell)] = "ИтогоПолуч"
        sheet["B"+str(5+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['totalReceived']
        sheet["A"+str(6+cursorCell)] = "ИтогоОплач"
        sheet["B"+str(6+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['totalDiscarded']
        sheet["A"+str(7+cursorCell)] = "КонОстаток"
        sheet["B"+str(7+cursorCell)] = dataPacket[tpm + cursorDataPacket][1][0]['endAmount']
        #print("len(dataPacket[tpm + cursorDataPacket]):" + str(len(dataPacket[tpm + cursorDataPacket])))
        for i in range(len(dataPacket[tpm + cursorDataPacket][0])):
            #print("i:" + str(2+cursorCell))
            sheet["D"+str(2+cursorCell)] = line + onlyfiles[tpm]
            sheet["E"+str(2+cursorCell)] = dataPacket[tpm + cursorDataPacket][0][i]['docType']
            sheet["F"+str(2+cursorCell)] = dataPacket[tpm + cursorDataPacket][0][i]['docDate']
            if dataPacket[tpm + cursorDataPacket][0][i]['receivedDate']:
                prihod = dataPacket[tpm + cursorDataPacket][0][i]['paySum']#.replace('.', ',')
                rashod = "0.00"
            if dataPacket[tpm + cursorDataPacket][0][i]['discardedDate']:
                rashod = dataPacket[tpm + cursorDataPacket][0][i]['paySum']#.replace('.', ',')   
                prihod = "0.00"
            if not "." in prihod:
                prihod = prihod +".00"
            #print(prihod)
            if not "." in rashod:
                rashod = rashod + ".00"
            #print(rashod)
            sheet["G"+str(2+cursorCell)].value = "=" + prihod
            sheet["H"+str(2+cursorCell)].value = "=" + rashod
            
            dateObj = datetime.strptime(dataPacket[tpm + cursorDataPacket][0][i]['docDate'], '%d.%m.%Y')             
            sheet["I"+str(2+cursorCell)].value = quarter(dateObj)
            sheet["J"+str(2+cursorCell)].value = month(dateObj)
                        
            if prihod != "0.00":
                middleVipiska.addValuePrihod(prihod,dateObj)
            if rashod != "0.00":
                middleVipiska.addValueRashod(rashod,dateObj)

            cursorCell += 1
        mv = middleVipiska.getTable()
        
        #печатаем график по выписке
        print(mv[0])
        print("---------------------------------")
        print(mv[1])
        print("---------------------------------")
        print(mv[2])
        cursorMiddle = 0
        sheet["L"+str(2+cursorCell+cursorMiddle-1)] = "Месяц"
        sheet["M"+str(2+cursorCell+cursorMiddle-1)] = "Приход"
        sheet["N"+str(2+cursorCell+cursorMiddle-1)] = "Расход"
        sheet["O"+str(2+cursorCell+cursorMiddle-1)] = "Баланс"
        for i in mv[0]:
            sheet["L"+str(2+cursorCell+cursorMiddle)] = i
            sheet["M"+str(2+cursorCell+cursorMiddle)] = mv[0][i]
            sheet["N"+str(2+cursorCell+cursorMiddle)] = mv[1][i]
            sheet["O"+str(2+cursorCell+cursorMiddle)] = mv[2][i]
            cursorMiddle += 1
        #строим график средний по месяцам
        min_row=2+cursorCell-1
        max_row=2+cursorCell+len(mv[0])-1
        values = Reference(sheet,
                   min_col=13,
                   max_col=15,
                   min_row=min_row,
                   max_row=max_row)
        chart = LineChart()
        chart.add_data(values, titles_from_data=True)
        chart.title = "Средняя оборотка по месяцам по выписке"
        dates = Reference(sheet, min_col=12, min_row=min_row+1, max_row=max_row)
        chart.set_categories(dates)

        chart.x_axis.title = "Месяцы"
        chart.y_axis.title = "Обороты"
        sheet.add_chart(chart,"R"+str(2+cursorCell))
        
        mv[0].clear()
        mv[1].clear()
        mv[2].clear()
        temp = list(mv)
        temp.clear()
        mv = tuple(temp)
        
    del middleVipiska
    #time.sleep(5)
    cursorDataPacket += len(onlyfiles)    
    #print(utils.FORMULAE)
    workbook.save(filename=line+"\\"+filename)


        