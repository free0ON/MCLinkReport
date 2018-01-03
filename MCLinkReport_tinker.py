'''
v1.7 на tkinter
Программа конвертации отчетов программы MCLink в формате xml в формат xls
Проблема с округлением
Убрал размерность
---
Автооткрытие
Автозапуск
Сохранение настроек
Ошибки в результатах
Протокол AsFound и AsReturn 
'''
import datetime
import os
import sys
import xml.etree.ElementTree as etree
from threading import Thread
from time import sleep
import configparser

import xlrd
import xlwt
from xlutils.copy import copy as xlcopy
from tkinter import *
from tkinter.filedialog import *

CSM = 'КрасЦСМ'

Title = CSM + ". Сохранение протоколов поверки MCLink v1.7"

class DemonConvertation(Thread):
    runing = bool
    pathname = str(os.path.dirname(sys.argv[0])).replace('/','\\')
    xml_folder = pathname
    Excel_folder = pathname
    template_filename = pathname + '\\шаблон.xls'
    autoopen = bool
    config_filename = 'config.ini'
    conf = configparser.RawConfigParser()

    def __init__(self):
        Thread.__init__(self)
        self.update_settings()

    def setXmlFolder(self, xml_folder):
        self.conf.read(self.config_filename)
        self.conf.set('path', 'xml', xml_folder)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.xml_folder = xml_folder

    def setExcelFolder(self, Excel_folder):
        self.conf.read(self.config_filename)
        self.conf.set('path', 'Excel', Excel_folder)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.Excel_folder = Excel_folder

    def setTemplateFilename(self, template_filename):
        self.conf.read(self.config_filename)
        self.conf.set('path', 'Template', template_filename)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.template_filename = template_filename

    def setAutoOpen(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto','autoopen',str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.autoopen = set

    def setAutoStart(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autostart', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.runing = set

    # Функция обновления настроек
    def update_settings(self):
        self.conf.read(self.config_filename)
        xml_folder = self.conf.get('path','xml')
        Excel_folder = self.conf.get('path','Excel')
        template_filename = self.conf.get('path','Template')
        autostart = self.conf.get('auto','autostart')
        autoopen = self.conf.get('auto','autoopen')
        if xml_folder != '':
            self.xml_folder = str(xml_folder)
        if Excel_folder != '':
            self.Excel_folder = str(Excel_folder)
        if template_filename != '':
            self.template_filename = str(template_filename)
        if autostart == 'True':
            self.runing = True
        else:
            self.runing = False
        if autoopen == 'True':
            self.autoopen = True
        else:
            self.autoopen = False

    def run(self):
        self.update_settings()
        while self.runing:
            file = os.listdir(self.xml_folder)
            if len(file) != 0:
                for i in file:
                    ext = i[-4:]
                    if ext== '.xml':
                        self.convertation(self.xml_folder + '\\' + i)
                        sleep(1)
            sleep(1)

    def roundStr(self, _str, num):
        _str = str(_str).replace(',','.')
        _str = float(_str)
        _str = round(_str,num)
        return str(_str).replace('.',',')

    def rightFileName(self, _str):
        _str = _str.replace('#', '',200)
        _str = _str.replace('&', '',200)
        _str = _str.replace(':', '',200)
        _str = _str.replace('<', '',200)
        _str = _str.replace('>', '',200)
        _str = _str.replace('?', '',200)
        _str = _str.replace('/', '',200)
        _str = _str.replace('\\', '',200)
        _str = _str.replace('\"', '',200)
        _str = _str.replace('|', '',200)
        _str = _str.replace('*', '',200)
        _str = _str.replace('&', '',200)
        _str = _str.replace('\n','',200)
        return _str

    def convertation(self, xml_filename=None):
        tree = etree.parse(xml_filename)
        root = tree.getroot()
        date_time = str(datetime.datetime.now()).replace(':', '')

        WeightSetCalibration = root.find('WeightSetCalibration')
        ContactOwner = WeightSetCalibration.find('ContactOwner')
        Company = str(ContactOwner.find('Company').text).strip(' ')
        Department = str(ContactOwner.find('Department').text).strip(' ')

        City = ContactOwner.find('City')

        TestWeightSet = WeightSetCalibration.find('TestWeightSet')
        TestWeightSet_Description = TestWeightSet.find('Description').text
        TestWeightCalibrations = TestWeightSet.find('TestWeightCalibrations')


        EnvironmentalConditions = WeightSetCalibration.find('EnvironmentalConditions')
        AirTemperature = EnvironmentalConditions.find('AirTemperature')
        AirPressure = EnvironmentalConditions.find('AirPressure')
        Humidity = EnvironmentalConditions.find('Humidity')
        AirDensity = EnvironmentalConditions.find('AirDensity')
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
        if Department == 0:
            Company_Name = Company  # назвение заказчика
        else:
            Company_Name = Company + ' ' + Department

        # Address = City.get('ZipCode') + ', ' + City.get('State') + ', ' + ContactOwner.find('Address').text  # адрес
        AirDensity_Min = self.roundStr(AirDensity.find('Min').text,4)
        AirDensity_Max = self.roundStr(AirDensity.find('Max').text,4)
        AirDensity_Avr = self.roundStr(AirDensity.find('Average').text,4)


        AirTemperature_Min = self.roundStr(AirTemperature.get('Min'),2)  # температура мин
        AirTemperature_Max = self.roundStr(AirTemperature.get('Max'),2)  # температура макс
        AirTemperature_Avr = self.roundStr(AirTemperature.get('Average'),2)  # температура средняя
        # AirTemperature_Unit = AirTemperature.get('Unit')  # размерность температуры

        AirPressure_Min = self.roundStr(AirPressure.get('Min'),2)  # давление мин
        AirPressure_Max = self.roundStr(AirPressure.get('Max'),2)  # давление макс
        AirPressure_Avr = self.roundStr(AirPressure.get('Average'), 2)  # давление среднее
        # AirPressure_Unit = AirPressure.get('Unit')  # размерность давления

        Humidity_Min = self.roundStr(Humidity.get('Min'),2)  # влажность мин
        Humidity_Max = self.roundStr(Humidity.get('Max'),2)  # влажность макс
        Humidity_Avr = self.roundStr(Humidity.get('Average'), 2)  # влажность средняя

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
        TestWeightCalibrationAsFound = TestWeightCalibrations.findall('TestWeightCalibrationAsFound')
        TestWeightSet_AlloyMaterials = TestWeightSet.find('AlloyMaterials')
        TestWeightSet_AlloyMaterial = TestWeightSet_AlloyMaterials.findall('AlloyMaterial')[0]
        Density = TestWeightSet_AlloyMaterial.get('Density') + TestWeightSet_AlloyMaterial.get('DensityUnit')
        Test_Passed = True

        # Есть отрицательные результаты или ошибочно записанные AsFound
        if len(TestWeightCalibrationAsFound) > 0:
            for found in TestWeightCalibrationAsFound:
                Nominal = float(str(found.find('Nominal').text).replace(',','.'))
                Error = float(str(found.find('ConventionalMassCorrection').text).replace(',','.'))
                Tolerance =  float(str(found.find('Tolerance').text).replace(',','.'))
                # Отрицательный результат
                if abs(Error) < 0.1*Nominal*1000 and abs(Error) > Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = False
                # Ошибочно записанный положительный результат
                elif abs(Error) <= Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = True

        # Есть положительные результаты AsReturned
        if len(TestWeightCalibrationAsReturned) > 0:

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

            # КрасЦСМ номер протокола не печатаем
            if CSM != "КрасЦСМ":
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

            ws.write(14, 2, Density)

            ws.write(31, 1, AirTemperature_Min, styleCellCenter)
            ws.write(32, 1, AirTemperature_Max, styleCellCenter)
            ws.write(33, 1, AirTemperature_Avr, styleCellCenter)

            ws.write(31, 3, Humidity_Min, styleCellCenter)
            ws.write(32, 3, Humidity_Max, styleCellCenter)
            ws.write(33, 3, Humidity_Avr, styleCellCenter)

            ws.write(31, 5, AirPressure_Min, styleCellCenter)
            ws.write(32, 5, AirPressure_Max, styleCellCenter)
            ws.write(33, 5, AirPressure_Avr, styleCellCenter)

            ws.write(31, 7, AirDensity_Min, styleCellCenter)
            ws.write(32, 7, AirDensity_Max, styleCellCenter)
            ws.write(33, 7, AirDensity_Avr, styleCellCenter)

            # TODO: НАстройки в шаблон
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
            # TODO: Настройки в шаблон
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

                Method = StepSeriesIndex[0:6]
                # ABA
                if Method == 'A1B1A1' or Method == '(A)1B1':  # 1 ABA
                    for cicle in range(int(len(WeightReading) / 3)):
                        for x in range(RowWeightReading, RowWeightReading + 3):
                            for y in range(2, 8):
                                ws.write(x, y, '', styleCellBorder)
                        A1.append(WeightReading[cicle*3].get('WeightReading'))
                        WeightReadingUnit = WeightReading[cicle].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle*3 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle*3 + 2].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A2 ' + A2[cicle],
                                 styleCellLeftSinglCell)
                        A1[cicle] = float(A1[cicle].replace(',', '.'))
                        B1[cicle] = float(B1[cicle].replace(',', '.'))
                        A2[cicle] = float(A2[cicle].replace(',', '.'))
                        Diff.append(round(B1[cicle] - (A1[cicle] + A2[cicle]) / 2, 4))
                        Avr += Diff[cicle]
                        ws.write(RowWeightReading, 2, str(Diff[cicle]).replace('.', ','),
                                 styleCellLeftBottom)
                        RowWeightReading += 1

                if Method == 'A1B1B1':  # 1 ABBA
                    for cicle in range(int(len(WeightReading) / 4)):
                        for x in range(RowWeightReading, RowWeightReading + 4):
                            for y in range(2, 8):
                                ws.write(x, y, '', styleCellBorder)
                        A1.append(WeightReading[cicle*4].get('WeightReading'))
                        WeightReadingUnit = WeightReading[0].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle*4 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B2.append(WeightReading[cicle*4 + 2].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B2 ' + B2[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle*4 + 3].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A2 ' + A2[cicle],
                                 styleCellLeftSinglCell)
                        A1[cicle] = float(A1[cicle].replace(',', '.'))
                        B1[cicle] = float(B1[cicle].replace(',', '.'))
                        B2[cicle] = float(B2[cicle].replace(',', '.'))
                        A2[cicle] = float(A2[cicle].replace(',', '.'))
                        Diff.append(round((B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2, 4))
                        Avr += Diff[cicle]
                        ws.write(RowWeightReading, 2, str(Diff[cicle]).replace('.', ','),
                                 styleCellLeftBottom)
                        RowWeightReading += 1

                Avr = str(round(Avr / len(Diff), 4)).replace('.', ',')

                ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
                ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection,4)
                ConventionalMassCorrectionUnit = i.find('ConventionalMassCorrectionUnit').text
                ConventionalMass = i.find('ConventionalMass').text
                #ConventionalMass = self.roundStr(ConventionalMass,4)
                ConventionalMassUnit = i.find('ConventionalMassUnit').text
                ExpandedMassUncertainty = i.find('ExpandedMassUncertainty').text
                ExpandedMassUncertaintyUnit = i.find('ExpandedMassUncertaintyUnit').text
                ws.write(Row, 3, Avr, styleCellCenterSinglCellTop)
                ReferenceWeight_ConventionalMassError = self.roundStr(ReferenceWeight_ConventionalMassError,4)
                ws.write(Row, 4, ReferenceWeight_ConventionalMassError,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 5, ConventionalMassCorrection,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 6, ConventionalMass, styleCellCenterSinglCellTop)
                ws.write(Row, 7, ExpandedMassUncertainty, styleCellCenterSinglCellTop)
                for x in range(Row + 1, Row + len(WeightReading)):
                    ws.write(x, 0, '', styleCellLeftLine)
                    ws.write(x, 8, '', styleCellRightLine)
                Row = RowWeightReading

            for y in range(0, 9):
                ws.write(Row, y, '', styleCellTopLine)

            if Test_Passed:
                ws.write(Row + 2, 0, 'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                ws.write(Row + 3, 0, 'На основании результатов поверки выдано свидетельство о поверке № _______________ от ___.__________.______г.')
            else:
                if CSM != 'КрасЦСМ':
                    ws.write(Row + 2, 0, 'Заключение по результатам поверки: гири не пригодны к использованию по классу точности '+TestWeightSet_AccuracyClass+' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 3, 0, 'На основании результатов поверки выдано извещение о непригодности № _______________ от ___.__________.______г.')
                else:
                    ws.write(Row + 2, 0,'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 3, 0,'На основании результатов поверки выдано свидетельство о поверке № _______________ от ___.__________.______г.')

                pass

            ws.write(Row + 6, 0, 'Поверитель:_____________________ ' + CalibratedBy)
            ws.write(Row + 6, 6, 'Дата протокола: ' + str(datetime.date.today().day) +'.' +str(datetime.date.today().month)+'.' +str(datetime.date.today().year))

            # сохранение данных в новый документ
            date_time = str(datetime.datetime.now()).replace(':', '')
            date_time = str(date_time).replace('.', '_')
            ws.insert_bitmap('logo.bmp',1,7)
            file_to_save = self.rightFileName(Company_Name + ' ' + TestWeightSet_AccuracyClass + ' '+ TestWeightSet_SerialNumber +' ' + date_time + '.xls')
            wb.save(self.Excel_folder + '\\'+file_to_save)
            if self.autoopen == True:
                os.startfile(self.Excel_folder + '\\' + file_to_save)
        # Удаляем исходный файл xml
        os.remove(xml_filename)

# класс главного окна
class MainForm:
    def selectXmlFolder(self):
        folder = askdirectory()
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setXmlFolder(folder)
            self.lbXmlFolder["text"] = folder

    def selectExcelFolder(self):
        folder = askdirectory()
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setExcelFolder(folder)
            self.lbExcelFolder["text"] = folder

    def close(self):
        self.demon.runing = False
        root.destroy()
        exit()

    def start(self):
        self.demon = DemonConvertation()
        self.demon.setAutoStart(True)
        self.demon.setDaemon(True)
        self.btStart['state'] = 'disabled'
        self.btStop['state'] = 'active'
        self.demon.start()

    def stop(self):
        self.demon.setAutoStart(False)
        self.btStart['state'] = 'active'
        self.btStop['state'] = 'disabled'

    def change(self):
        if self.autoopen.get() == 1:
            self.demon.setAutoOpen(True)
        else:
            self.demon.setAutoOpen(False)

    def __init__(self, master):
        self.demon = DemonConvertation()
        self.imgExit = Image
        self.master = master
        self.master.title(Title)
        self.master.geometry('600x200')
        self.imgExit = PhotoImage(file="icons\\exit.png")
        self.btQuit = Button(self.master, image = self.imgExit, command = self.close)
        self.btQuit.place(x=10, y=10, width=30, height=30)

        self.imgStart = PhotoImage(file="icons\\player_play.png")
        self.btStart = Button(self.master, image=self.imgStart, command = self.start)
        self.btStart.place(x=40, y=10, width=30, height=30)
        self.imgStop = PhotoImage(file="icons\\player_stop.png")
        self.btStop = Button(self.master, image=self.imgStop, command = self.stop, state = 'disabled')
        self.btStop.place(x=70, y=10, width=30, height=30)

        self.lbXmlFolder = Label(self.master, text='Папка xml: ' + self.demon.xml_folder)
        self.lbXmlFolder.place(x=10, y=50, width=500, height=20)

        self.lbExcelFolder = Label(self.master, text='Папка Excel: ' + self.demon.Excel_folder)
        self.lbExcelFolder.place(x=10, y=100, width=500, height=20)

        self.btXmlFolder = Button(self.master, text='...', command = self.selectXmlFolder)
        self.btXmlFolder.place(x=550, y=50,width=30, height=20)
        self.btExcelFolder = Button(self.master, text='...',command = self.selectExcelFolder)
        self.btExcelFolder.place(x=550, y=100, width=30, height=20)
        self.autoopen = IntVar()

        self.chbAutoOpen = Checkbutton(self.master, text='Автооткрытие протокола', variable=self.autoopen, onvalue=1, offvalue=0, command =self.change)
        if self.demon.autoopen == True:
            self.chbAutoOpen.select()
        self.chbAutoOpen.place(x=10, y=150, width=200, height=30)

        if self.demon.runing == True:
            self.btStart['state'] = 'disabled'
            self.btStop['state'] = 'active'
            self.demon.start()

        self.master.protocol('WM_DELETE_WINDOW',self.close)
        self.master.mainloop()

# создание окна
root = Tk()
# root.protocol('WM_DELETE_WINDOW',MainForm.close())

# запуск окна
MainForm(root)