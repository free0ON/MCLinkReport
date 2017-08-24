'''
Программа конвертации отчетов программы MCLink в формате xml в формат xls

'''
import datetime
import os
import sys
import xml.etree.ElementTree as etree

from threading import Thread
from time import sleep


import xlrd
import xlwt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QPushButton, QAction, QApplication, QFileDialog, QLabel, QCheckBox, QToolBar
from xlutils.copy import copy as xlcopy


class DemonConvertation(Thread):
    runing = False
    pathname = str(os.path.dirname(sys.argv[0])).replace('/','\\')
    scan_folder = pathname
    dest_folder = pathname
    template_filename = 'шаблон_новый.xls'
    autoopen = False

    def __init__(self):
        Thread.__init__(self)

    def run(self):
        while self.runing:
            file = os.listdir(self.scan_folder)
            if len(file) != 0:
                for i in file:
                    ext = i[-4:]
                    if ext== '.xml':
                        self.convertation(self.scan_folder + '\\' + i)
                        sleep(1)
            sleep(1)

    def convertation(self, xml_filename=None):
        tree = etree.parse(xml_filename)
        root = tree.getroot()
        date_time = str(datetime.datetime.now()).replace(':', '')

        WeightSetCalibration = root.find('WeightSetCalibration')
        ContactOwner = WeightSetCalibration.find('ContactOwner')
        Company = ContactOwner.find('Company')
        City = ContactOwner.find('City')

        TestWeightSet = WeightSetCalibration.find('TestWeightSet')
        TestWeightSet_Description = TestWeightSet.find('Description').text
        TestWeightCalibrations = TestWeightSet.find('TestWeightCalibrations')


        EnvironmentalConditions = WeightSetCalibration.find('EnvironmentalConditions')
        AirTemperature = EnvironmentalConditions.find('AirTemperature')
        AirPressure = EnvironmentalConditions.find('AirPressure')
        Humidity = EnvironmentalConditions.find('Humidity')

        Methods = WeightSetCalibration.find('Methods')
        Method = Methods.findall('Method')

        ReferenceWeightSets = WeightSetCalibration.find('ReferenceWeightSets')
        ReferenceWeightSet = ReferenceWeightSets.findall('ReferenceWeightSet')

        ReferenceWeights = WeightSetCalibration.find('ReferenceWeights')
        ReferenceWeight = ReferenceWeights.findall('ReferenceWeight')

        MassComparators = WeightSetCalibration.find('MassComparators')
        MassComparator = MassComparators.findall('MassComparator')

        TestWeightCalibrations_Count = TestWeightCalibrations.get('Count')  # количество гирь
        EndDate = WeightSetCalibration.get('EndDate')  # дата поверки
        CertificateNumber = WeightSetCalibration.get('CertificateNumber')  # номер сертификата
        CalibratedBy = WeightSetCalibration.get('CalibratedBy')  # поверитель
        CustomerNumber = ContactOwner.get('CustomerNumber')  # ИНН
        Company_Name = Company.text  # назвение заказчика
        # Address = City.get('ZipCode') + ', ' + City.get('State') + ', ' + ContactOwner.find('Address').text  # адрес
        AirTemperature_Min = AirTemperature.get('Min')  # температура мин
        AirTemperature_Max = AirTemperature.get('Max')  # температура макс
        AirTemperature_Avr = AirTemperature.get('Average')  # температура средняя
        # AirTemperature_Unit = AirTemperature.get('Unit')  # размерность температуры

        AirPressure_Min = AirPressure.get('Min')  # давление мин
        AirPressure_Max = AirPressure.get('Max')  # давление макс
        AirPressure_Avr = AirPressure.get('Average')  # давление среднее
        # AirPressure_Unit = AirPressure.get('Unit')  # размерность давления

        Humidity_Min = Humidity.get('Min')  # влажность мин
        Humidity_Max = Humidity.get('Max')  # влажность макс
        Humidity_Avr = Humidity.get('Average')  # влажность средняя
        # Humidity_Unit = Humidity.get('Unit')  # размерность влажности

        # Method_Name = Method[0].get('Name')  # метод поверки
        # Method_Process = Method[0].get('Process')  # название процесса поверки

        MassComparator_Model = MassComparator[0].get('Model')  # модель компаратора
        MassComparator_SerialNumber = MassComparator[0].get('SerialNumber')  # серийный номер компаратора
        MassComparator_Description = MassComparator[0].find(
            'Description').text  # описание компаратора (дискретность, ско...)

        ReferenceWeightSet_SerialNumber = ReferenceWeightSet[0].get(
            'SerialNumber')  # серийный номер набора эталонов
        ReferenceWeightSet_Class = ReferenceWeightSet[0].get('Class')  # класс набора эталонов
        # ReferenceWeightSet_NextCalibrationDate = ReferenceWeightSet[0].get('NextCalibrationDate')  # дата следующей калибровки эталонов

        # ReferenceWeight_Array = []  # массив наборов эталонов
        TestWeightSet_SerialNumber = TestWeightSet.get('SerialNumber')
        TestWeightSet_AccuracyClass = TestWeightSet.get('AccuracyClass')
        TestWeightSet_Manufacturer =  TestWeightSet.get('Manufacturer')
        TestWeightSet_Range = TestWeightSet.get('Range')
        TestWeightCalibrationAsReturned = TestWeightCalibrations.findall('TestWeightCalibrationAsReturned')

        rb = xlrd.open_workbook(self.template_filename, formatting_info=True, on_demand=True)  # открываем книгу
        # rs = rb.get_sheet(0)  # выбираем лист ?

        wb = xlcopy(rb)  # копируем книгу в память
        ws = wb.get_sheet(0)  # выбираем лист

        # стиль ячеки выравнивание по центру
        styleCellCenter = xlwt.easyxf('border: top thin, left thin, bottom thin, right thin; align: horiz center')
        # стиль ячейки выравнивание влево
        styleCellLeft = xlwt.easyxf('border: top thin, left thin, bottom thin; align: horiz left')
        styleCellLeftSinglCell = xlwt.easyxf(
            'border: top thin, left thin, bottom thin, right thin; align: horiz left')
        # styleCellCenterSinglCell = xlwt.easyxf('border: top thin, left thin, bottom thin, right thin; align: horiz center')
        styleCellCenterSinglCellTop = xlwt.easyxf('border: top thin, left thin, right thin; align: horiz center')
        styleCellTopLine = xlwt.easyxf('border: top thin')
        styleCellLeftLine = xlwt.easyxf('border: left thin')
        styleCellRightLine = xlwt.easyxf('border: right thin')

        styleCellBorder = xlwt.easyxf('border: left thin, right thin')

        styleCellLeftBottom = xlwt.easyxf('border: left thin, bottom thin, right thin; align: horiz left')

        if TestWeightCalibrations_Count == '1':
            CI_Name = 'Гиря'
        else:
            CI_Name = 'Набор гирь'

        ws.write(1, 4, CertificateNumber)  # номер протокола
        ws.write(2, 1, EndDate)  # дата поверки
        ws.write(3, 1, CI_Name)  # наименование СИ
        ws.write(2, 6, TestWeightSet_AccuracyClass)  # класс точности
        ws.write(3, 6, TestWeightSet_Range) # номинальное заначение массы
        ws.write(4, 6, TestWeightSet_SerialNumber)  # серийный номер
        ws.write(4, 1, TestWeightSet_Description)  # год выпуска
        ws.write(5, 1, Company_Name)  # название заказчика
        ws.write(6, 1, CustomerNumber)  # номер заказчика
        ws.write(7, 1, TestWeightSet_Manufacturer) # производитель гирь


        ws.write(31, 2, AirTemperature_Min, styleCellCenter)
        ws.write(32, 2, AirTemperature_Max, styleCellCenter)
        ws.write(33, 2, AirTemperature_Avr, styleCellCenter)

        ws.write(31, 4, Humidity_Min, styleCellCenter)
        ws.write(32, 4, Humidity_Max, styleCellCenter)
        ws.write(33, 4, Humidity_Avr, styleCellCenter)

        ws.write(31, 6, AirPressure_Min, styleCellCenter)
        ws.write(32, 6, AirPressure_Max, styleCellCenter)
        ws.write(33, 6, AirPressure_Avr, styleCellCenter)

        row = 37
        for ref in ReferenceWeightSet:
            # название набора гирь
            ws.write(row, 0, 'Набор гирь', styleCellCenter)
            ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
            ws.write(row, 2, ReferenceWeightSet_SerialNumber, styleCellCenter)
            ReferenceWeightSet_Class = ref.get('Class')
            # класс точности набора
            ws.write(row, 4, ReferenceWeightSet_Class, styleCellCenter)
            row += 1

        for comp in MassComparator:
            # название компаратора
            MassComparator_Model = comp.get('Model')
            ws.write(row, 0, MassComparator_Model,styleCellCenter)
            MassComparator_SerialNumber = comp.get('SerialNumber')
            ws.write(row, 2, MassComparator_SerialNumber, styleCellCenter)
            MassComparator_Description = comp.find('Description').text
            # описание компаратора. В поле Описание (Description) должны быть записаны дискретность и СКО модели компаратора
            ws.write(row, 4, MassComparator_Description, styleCellCenter)
            row += 1

        Row = 46
        for i in TestWeightCalibrationAsReturned:
            Nominal = i.find('Nominal').text
            NominalUnit = i.find('NominalUnit').text
            WeightID = i.find('WeightID').text
            Index = i.find('ReferenceWeight').text
            ReferenceWeight_ConventionalMassError = 0
            Tolerance = i.find('Tolerance').text
            for j in ReferenceWeight:
                if Index == j.get('Index'):
                    ReferenceWeight_ConventionalMassError = j.get('ConventionalMassError')
            if TestWeightCalibrations_Count == '1':
                ws.write(Row, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
            else:
                ws.write(Row, 0, str.strip(WeightID + Nominal + NominalUnit), styleCellCenterSinglCellTop)
            ws.write(Row, 8, str(Tolerance), styleCellCenterSinglCellTop)
            MeasurementReadings = i.find('MeasurementReadings')
            WeightReading = MeasurementReadings.findall('WeightReading')
            RowWeightReading = Row

            A1 = []
            A2 = []
            B1 = []
            B2 = []
            Diff = []
            Avr = 0
            WeightReadingUnit = 0

            # Определение метода
            Method = ''
            StepSeriesIndex = ''
            for wr in WeightReading:
                StepSeriesIndex += wr.get('Step') + wr.get('SeriesIndex')
            if len(StepSeriesIndex) > 3:
                # ABBA
                if StepSeriesIndex[0:3] == 'A1B1B1A1':
                    Method = 'ABBA'    
                
                # ABA
                if StepSeriesIndex[0:3] == 'A1B1A1A2':
                    Method = 'ABA'

                # ABABA
                if StepSeriesIndex[0:3] == 'A1B1A2B2':
                    Method = 'ABABA'
            if len(StepSeriesIndex) == 3:
                # ABA
                if StepSeriesIndex[0:2] == 'A1B1A1':
                    Method = 'ABA'


            # ABA
            if ((len(WeightReading) % 3) == 0 ):  # 1 ABA
                for cicle in range(int(len(WeightReading) / 3)):
                    for x in range(RowWeightReading, RowWeightReading + 3):
                        for y in range(2, 8):
                            ws.write(x, y, '', styleCellBorder)
                    A1.append(WeightReading[cicle].get('WeightReading'))
                    WeightReadingUnit = WeightReading[cicle].get('WeightReadingUnit')
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    RowWeightReading += 1
                    B1.append(WeightReading[cicle + 1].get('WeightReading'))
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    RowWeightReading += 1
                    A2.append(WeightReading[cicle + 2].get('WeightReading'))
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' A2 ' + A2[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    A1[cicle] = float(A1[cicle].replace(',', '.'))
                    B1[cicle] = float(B1[cicle].replace(',', '.'))
                    A2[cicle] = float(A2[cicle].replace(',', '.'))
                    Diff.append(round(B1[cicle] - (A1[cicle] + A2[cicle]) / 2, 4))
                    Avr += Diff[cicle]
                    ws.write(RowWeightReading, 2, str(Diff[cicle]).replace('.', ',') + WeightReadingUnit,
                             styleCellLeftBottom)
                    RowWeightReading += 1

            if ((len(WeightReading) % 4) == 0):  # 1 ABBA
                for cicle in range(int(len(WeightReading) / 4)):
                    for x in range(RowWeightReading, RowWeightReading + 4):
                        for y in range(2, 8):
                            ws.write(x, y, '', styleCellBorder)
                    A1.append(WeightReading[cicle].get('WeightReading'))
                    WeightReadingUnit = WeightReading[0].get('WeightReadingUnit')
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    RowWeightReading += 1
                    B1.append(WeightReading[cicle + 1].get('WeightReading'))
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    RowWeightReading += 1
                    B2.append(WeightReading[cicle + 2].get('WeightReading'))
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' B2 ' + B2[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    RowWeightReading += 1
                    A2.append(WeightReading[cicle + 3].get('WeightReading'))
                    ws.write(RowWeightReading, 1, str(cicle + 1) + ' A2 ' + A2[cicle] + WeightReadingUnit,
                             styleCellLeftSinglCell)
                    A1[cicle] = float(A1[cicle].replace(',', '.'))
                    B1[cicle] = float(B1[cicle].replace(',', '.'))
                    B2[cicle] = float(B2[cicle].replace(',', '.'))
                    A2[cicle] = float(A2[cicle].replace(',', '.'))
                    Diff.append(round((B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2, 4))
                    Avr += Diff[cicle]
                    ws.write(RowWeightReading, 2, str(Diff[cicle]).replace('.', ',') + WeightReadingUnit,
                             styleCellLeftBottom)
                    RowWeightReading += 1

            Avr = str(round(Avr / len(Diff), 4)).replace('.', ',')

            ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
            ConventionalMassCorrectionUnit = i.find('ConventionalMassCorrectionUnit').text
            ConventionalMass = i.find('ConventionalMass').text
            ConventionalMassUnit = i.find('ConventionalMassUnit').text
            ExpandedMassUncertainty = i.find('ExpandedMassUncertainty').text
            ExpandedMassUncertaintyUnit = i.find('ExpandedMassUncertaintyUnit').text
            ws.write(Row, 3, Avr + WeightReadingUnit, styleCellCenterSinglCellTop)
            ws.write(Row, 4, ReferenceWeight_ConventionalMassError + ConventionalMassCorrectionUnit,
                     styleCellCenterSinglCellTop)
            ws.write(Row, 5, ConventionalMassCorrection + ConventionalMassCorrectionUnit,
                     styleCellCenterSinglCellTop)
            ws.write(Row, 6, ConventionalMass + ConventionalMassUnit, styleCellCenterSinglCellTop)
            ws.write(Row, 7, ExpandedMassUncertainty + ExpandedMassUncertaintyUnit, styleCellCenterSinglCellTop)
            for x in range(Row + 1, Row + len(WeightReading)):
                ws.write(x, 0, '', styleCellLeftLine)
                ws.write(x, 8, '', styleCellRightLine)
            Row = RowWeightReading

        for y in range(0, 9):
            ws.write(Row, y, '', styleCellTopLine)

        ws.write(Row + 2, 0, 'Заключение по результатам поверки:')
        ws.write(Row + 3, 0, 'На основании результатов поверки выдано свидетельство о поверке №')

        ws.write(Row + 6, 0, 'Поверитель:_____________________ ' + CalibratedBy)
        ws.write(Row + 6, 6, 'Дата протокола: ' + str(datetime.date.today().day) +'.' +str(datetime.date.today().month)+'.' +str(datetime.date.today().year))

        # сохранение данных в новый документ
        date_time = str(datetime.datetime.now()).replace(':', '')
        ws.insert_bitmap('logo.bmp',1,7)

        wb.save(self.dest_folder + '\\' + date_time + '.xls')
        os.remove(xml_filename)
        if self.autoopen == True:
            os.startfile(self.dest_folder + '\\' + date_time + '.xls')

class MainWindow(QMainWindow):
    demon = DemonConvertation
    startAction = QAction
    stopAction = QAction
    exitAction = QAction
    toolbar = QToolBar
    label1 = QLabel
    label2 = QLabel
    label3 = QLabel
    lbScanFolder = QLabel
    lbDestFolder = QLabel
    lbTemplate = QLabel
    btScanFolder = QPushButton
    btDestFolder = QPushButton
    btTemplate = QPushButton
    chbAutoOpen = QCheckBox

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.exitAction = QAction(QIcon('icons\\exit.png'), 'Выход', self)
        self.exitAction.setShortcut('Ctrl+Q')
        self.exitAction.setStatusTip('Выход')
        self.exitAction.triggered.connect(self.close)

        self.startAction = QAction(QIcon('icons\\player_play.png'), 'Запуск', self)
        self.startAction.setShortcut('Ctrl+C')
        self.startAction.setStatusTip('Запуск автосохранения')
        self.startAction.triggered.connect(self.start)
        self.startAction.setVisible(True)

        self.stopAction = QAction(QIcon('icons\\player_stop.png'), 'Остановка', self)
        self.stopAction.setShortcut('Ctrl+S')
        self.stopAction.setStatusTip('Остановка автосохранения')
        self.stopAction.triggered.connect(self.stop)
        self.stopAction.setVisible(False)

        self.toolbar = self.addToolBar('Tools')
        self.toolbar.addAction(self.exitAction)
        self.toolbar.addAction(self.startAction)
        self.toolbar.addAction(self.stopAction)

        self.label1 = QLabel('Папка xml:', self)
        self.label1.move(10, 50)
        self.label1.resize(60, 20)
        self.lbScanFolder = QLabel(self.demon.scan_folder, self)
        self.lbScanFolder.move(80, 50)
        self.lbScanFolder.resize(500, 20)
        self.btScanFolder = QPushButton("...", self)
        self.btScanFolder.move(550, 50)
        self.btScanFolder.resize(30, 20)
        self.btScanFolder.clicked.connect(self.selectScanFolder)

        self.label2 = QLabel('Папка Excel:', self)
        self.label2.move(10, 100)
        self.label2.resize(60, 20)
        self.lbDestFolder = QLabel(self.demon.dest_folder, self)
        self.lbDestFolder.move(80, 100)
        self.lbDestFolder.resize(500, 20)
        self.btDestFolder = QPushButton("...", self)
        self.btDestFolder.move(550, 100)
        self.btDestFolder.resize(30, 20)
        self.btDestFolder.clicked.connect(self.selectDestFolder)

        self.label3 = QLabel('Шаблон:',self)
        self.label3.move(10,150)
        self.label3.resize(60,20)
        self.lbTemplate = QLabel(self.demon.template_filename, self)
        self.lbTemplate.move(80,150)
        self.lbTemplate.resize(500,20)
        self.btTemplate = QPushButton('...',self)
        self.btTemplate.move(550,150)
        self.btTemplate.resize(30,20)
        self.btTemplate.clicked.connect(self.selectTemplate)

        self.chbAutoOpen = QCheckBox('Автооткрытие протокола', self)
        self.chbAutoOpen.move(10, 200)
        self.chbAutoOpen.resize(200, 30)
        self.chbAutoOpen.clicked.connect(self.changeAutoOpen)
        self.statusBar()
        self.setGeometry(500, 300, 600, 250)
        self.setWindowTitle('Сохранение протоколов поверки')
        self.setWindowIcon(QIcon('icons\\fileopen.png'))
        self.show()

    def start(self):
        self.startAction.setVisible(False)
        self.stopAction.setVisible(True)
        self.demon = DemonConvertation()
        self.demon.runing = True
        self.demon.setDaemon(True)
        self.demon.start()

    def stop(self):
        self.startAction.setVisible(True)
        self.stopAction.setVisible(False)
        self.demon.runing = False

    def selectScanFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите исходный файл', '')
        if folder != '':
            self.demon.scan_folder = str(folder).replace('/', '\\')
            self.lbScanFolder.setText(self.demon.scan_folder)

    def selectDestFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите папку сохранения отчетов', '')
        if folder != '':
            self.demon.dest_folder = str(folder).replace('/', '\\')
            self.lbDestFolder.setText(self.demon.dest_folder)

    def selectTemplate(self):
        template,ext = QFileDialog.getOpenFileName(self,'Выберите файл шаблона', self.demon.template_filename, '*.xls')
        if template != '':
            self.demon.template_filename = str(template).replace('/', '\\')
            self.lbTemplate.setText(self.demon.template_filename)

    def changeAutoOpen(self):
        if self.chbAutoOpen.isChecked() == True:
            self.demon.autoopen = True
        else:
            self.demon.autoopen = False


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
