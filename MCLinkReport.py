"""
Программа конвертации отчетов программы MCLink в формате xml в формат xls
v 2.4 Ошибки в формировании сертификатов и свидетельств
v 2.3 Метрологические характеистики на одном листе

v 2.1
Отключение протоколов галками
Коментарий набора в начало названия файла, если отключено свидетельство о поверке
--------
v2.0 QT на основе ui из QT designer
Переработан пользовательский интерфейс
Автооткрытие
Автозапуск
Сохранение настроек
"""

from datetime import datetime
import os
import sys
import xml.etree.ElementTree as etree
from threading import Thread
from time import sleep
import configparser

import xlrd
import xlwt
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from xlutils.copy import copy as xlcopy
from mainwindow import *
from dateutil.relativedelta import relativedelta

ver = "2.4"


# класс конвертора
class DemonConvertation(Thread):
    runned = bool
    pathname = str(os.path.dirname(sys.argv[0])).replace('/', '\\')
    xml_folder = pathname
    Excel_folder = pathname
    template_filename = pathname + '\\шаблон.xls'
    TemplateApprovalProtocol = ""
    TemplateApprovalCert = ""
    TemplateError = ""
    TemplateProtocolCal = ""
    TemplateCalCert = ""

    autoopen = bool
    autodelXML = bool
    ApprovalProtocolNum = 0
    ApprovalCertNum = 0
    CalCertNum = 0
    ErrorNum = 0
    CalProtocolNum = 0

    ApprovalProtocolEnable = bool
    ApprovalCertEnable = bool
    ErrorEnable = bool
    CalProtocolEnable = bool
    CalCertEnable = bool

    config_filename = 'config.ini'
    conf = configparser.RawConfigParser()
    CSM = ''
    ApprovalProtocolFolder = 'Протоколы поверки'
    ApprovalCertFolder = 'Свидетельства о поверке'
    ErrorFolder = 'Извещения о непригодности'
    CalProtocolFolder = 'Протоколы калибровки'
    CalCertFolder = 'Сертификаты о калибровке'

    def __init__(self):
        Thread.__init__(self)
        self.update_settings()

    def setProtocol(self, protocol, set):
        self.conf.read(self.config_filename)
        if protocol == 1:
            self.conf.set('enable', 'approvalprotocol', set)
            self.ApprovalProtocolEnable = set
        if protocol == 2:
            self.conf.set('enable', 'approvalsert', set)
            self.ApprovalCertEnable = set
        if protocol == 3:
            self.conf.set('enable', 'error', set)
            self.ErrorEnable = set
        if protocol == 4:
            self.conf.set('enable', 'calprotocol', set)
            self.CalProtocolEnable = set
        if protocol == 5:
            self.conf.set('enable', 'calsert', set)
            self.CalCertEnable = set

        with open(self.config_filename, "w") as config:
            self.conf.write(config)

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
        self.folderExist()

    def folderExist(self):
        if not os.access(self.Excel_folder + '\\' + self.ApprovalProtocolFolder, os.F_OK):
            os.mkdir(self.Excel_folder + '\\' + self.ApprovalProtocolFolder)
        if not os.access(self.Excel_folder + '\\' + self.ApprovalCertFolder, os.F_OK):
            os.mkdir(self.Excel_folder + '\\' + self.ApprovalCertFolder)
        if not os.access(self.Excel_folder + '\\' + self.CalProtocolFolder, os.F_OK):
            os.mkdir(self.Excel_folder + '\\' + self.CalProtocolFolder)
        if not os.access(self.Excel_folder + '\\' + self.CalCertFolder, os.F_OK):
            os.mkdir(self.Excel_folder + '\\' + self.CalCertFolder)
        if not os.access(self.Excel_folder + '\\' + self.ErrorFolder, os.F_OK):
            os.mkdir(self.Excel_folder + '\\' + self.ErrorFolder)

    def setTemplateFilename(self, template_filename, template):
        self.conf.read(self.config_filename)
        self.conf.set('path', template, template_filename)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        if template == 'TemplateApprovalProtocol': self.TemplateApprovalProtocol = template_filename
        if template == 'TemplateApprovalCert': self.TemplateApprovalCert = template_filename
        if template == 'TemplateCalProtocol': self.TemplateCalProtocol = template_filename
        if template == 'TemplateCalCert':  self.TemplateCalCert = template_filename
        if template == 'TemplateError':  self.TemplateError = template_filename

    def setAutoOpen(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autoopen', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.autoopen = set

    def setAutoStart(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autostart', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.runned = set

    def setAutoDelXML(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autodelXML', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.runned = set

    def setNameCSM(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('name', 'CSMName', str(set))
        with open(self.config_filename, 'w') as config:
            self.conf.write(config)
        self.CSM = set

    def setNums(self, set, name):
        self.conf.read(self.config_filename)
        self.conf.set('numdocs', name, str(set))
        with open(self.config_filename, 'w') as config:
            self.conf.write(config)

        if (name == 'ApprovalProtocolNum'):  self.ApprovalProtocolNum = set
        if (name == 'ApprovalCertNum'):  self.ApprovalCertNum = set
        if (name == 'CalCertNum'):  self.CalCertNum = set
        if (name == 'ErrorNum'):  self.ErrorNum = set
        if (name == 'CalProtocolNum'):  self.CalProtocolNum = set

    # Функция обновления настроек
    def update_settings(self):
        self.conf.read(self.config_filename)
        xml_folder = self.conf.get('path', 'xml')
        Excel_folder = self.conf.get('path', 'Excel')
        template_filename = self.conf.get('path', 'Template')
        autostart = self.conf.get('auto', 'autostart')
        autoopen = self.conf.get('auto', 'autoopen')
        autodelXML = self.conf.get('auto', 'autodelXML')
        CSM = self.conf.get('name', 'CSMName')

        if self.conf.get('enable', 'approvalprotocol') == 'True':
            self.ApprovalProtocolEnable = True
        else:
            self.ApprovalProtocolEnable = False

        if self.conf.get('enable', 'approvalsert') == 'True':
            self.ApprovalCertEnable = True
        else:
            self.ApprovalCertEnable = False

        if self.conf.get('enable', 'error') == 'True':
            self.ErrorEnable = True
        else:
            self.ErrorEnable = False

        if self.conf.get('enable', 'calprotocol') == 'True':
            self.CalProtocolEnable = True
        else:
            self.CalProtocolEnable = False

        if self.conf.get('enable', 'calsert') == 'True':
            self.CalCertEnable = True
        else:
            self.CalCertEnable = False

        self.ApprovalProtocolNum = self.conf.get('numdocs', 'ApprovalProtocolNum')
        self.ApprovalCertNum = self.conf.get('numdocs', 'ApprovalCertNum')
        self.CalCertNum = self.conf.get('numdocs', 'CalCertNum')
        self.ErrorNum = self.conf.get('numdocs', 'ErrorNum')
        self.CalProtocolNum = self.conf.get('numdocs', 'CalProtocolNum')

        self.TemplateApprovalProtocol = self.conf.get('path', 'TemplateApprovalProtocol')
        self.TemplateApprovalCert = self.conf.get('path', 'TemplateApprovalCert')
        self.TemplateCalProtocol = self.conf.get('path', 'TemplateCalProtocol')
        self.TemplateCalCert = self.conf.get('path', 'TemplateCalCert')
        self.TemplateError = self.conf.get('path', 'TemplateError')

        # startpath = os.path.dirname(sys.argv[0]).replace('/','\\')
        if os.access(xml_folder, 0) != True:
            self.setXmlFolder(str(self.pathname) + "\\xml")

        if os.access(Excel_folder, 0) != True:
            self.setExcelFolder(str(self.pathname) + "\\Excel")

        if os.access(self.TemplateApprovalProtocol, 0) != True:
            self.setTemplateFilename(str(self.pathname) + "\\templates\\Протокол поверки.xls",
                                     "TemplateApprovalProtocol")

        if os.access(self.TemplateApprovalCert, 0) != True:
            self.setTemplateFilename(str(self.pathname) + "\\templates\\Свидетельство о поверке.xls",
                                     "TemplateApprovalCert")

        if os.access(self.TemplateError, 0) != True:
            self.setTemplateFilename(str(self.pathname) + "\\templates\\Извещение о непригодности.xls", "TemplateError")

        if os.access(self.TemplateCalProtocol, 0) != True:
            self.setTemplateFilename(str(self.pathname) + "\\templates\\Протокол калибровки.xls", "TemplateCalProtocol")

        if os.access(self.TemplateCalCert, 0) != True:
            self.setTemplateFilename(str(self.pathname) + "\\templates\\Сертификат о калибровке.xls", "TemplateCalCert")

        if CSM != '':
            self.CSM = str(CSM)
        if xml_folder != '':
            self.xml_folder = str(xml_folder)
        if Excel_folder != '':
            self.Excel_folder = str(Excel_folder)
        if template_filename != '':
            self.template_filename = str(template_filename)
        if autostart == 'True':
            self.runned = True
        else:
            self.runned = False
        if autoopen == 'True':
            self.autoopen = True
        else:
            self.autoopen = False

    def run(self):
        self.update_settings()
        self.folderExist()
        while self.runned:
            file = os.listdir(self.xml_folder)
            if len(file) != 0:
                for i in file:
                    ext = i[-4:]
                    if ext == '.xml':
                        self.convertation(self.xml_folder + '\\' + i)
                        sleep(1)
            sleep(1)

    # округление строковых чисел
    def roundStr(self, _str, num):
        _str = str(_str).replace(',', '.')
        _str = float(_str)
        _str = round(_str, num)
        return str(_str).replace('.', ',')

    # проверка имени файла
    def rightFileName(self, _str):
        _str = _str.replace('#', '', 200)
        _str = _str.replace('&', '', 200)
        _str = _str.replace(':', '', 200)
        _str = _str.replace('<', '', 200)
        _str = _str.replace('>', '', 200)
        _str = _str.replace('?', '', 200)
        _str = _str.replace('/', '', 200)
        _str = _str.replace('\\', '', 200)
        _str = _str.replace('\"', '', 200)
        _str = _str.replace('|', '', 200)
        _str = _str.replace('*', '', 200)
        _str = _str.replace('&', '', 200)
        _str = _str.replace('\n', '', 200)
        _str = _str.replace(',', ' ', 200)
        return _str.strip()

    # формирование поверочного протокола
    def ApprovalReport(self, xml_filename):
        tree = etree.parse(xml_filename)
        root = tree.getroot()

        WeightSetCalibration = root.find('WeightSetCalibration')
        ContactOwner = WeightSetCalibration.find('ContactOwner')
        Company = str(ContactOwner.find('Company').text).strip(' ')
        Department = str(ContactOwner.find('Department').text).strip(' ')

        City = ContactOwner.find('City')

        TestWeightSet = WeightSetCalibration.find('TestWeightSet')
        TestWeightSet_Description = TestWeightSet.find('Description').text  # год выпуста
        TestWeightSet_Comment = TestWeightSet.find('Comment').text  # номер в ГР
        TestWeightSet_SerialNumber = TestWeightSet.get('SerialNumber')
        TestWeightSet_AccuracyClass = TestWeightSet.get('AccuracyClass')
        TestWeightSet_Manufacturer = TestWeightSet.get('Manufacturer')
        TestWeightSet_InternalID = TestWeightSet.get('InternalID')  # номер клейма предыдущей поверки
        TestWeightSet_Range = TestWeightSet.get('Range')

        str(TestWeightSet_Range).replace('g', 'г')
        str(TestWeightSet_Range).replace('kg', 'кг')
        str(TestWeightSet_Range).replace('mg', 'мг')

        TestWeightCalibrations = TestWeightSet.find('TestWeightCalibrations')
        TestWeightCalibrationAsReturned = TestWeightCalibrations.findall('TestWeightCalibrationAsReturned')
        TestWeightCalibrationAsFound = TestWeightCalibrations.findall('TestWeightCalibrationAsFound')
        TestWeightSet_AlloyMaterials = TestWeightSet.find('AlloyMaterials')
        TestWeightSet_AlloyMaterial = TestWeightSet_AlloyMaterials.findall('AlloyMaterial')[0]
        Density = TestWeightSet_AlloyMaterial.get('Density') + TestWeightSet_AlloyMaterial.get('DensityUnit')

        EnvironmentalConditions = WeightSetCalibration.find('EnvironmentalConditions')
        AirTemperature = EnvironmentalConditions.find('AirTemperature')
        AirPressure = EnvironmentalConditions.find('AirPressure')
        Humidity = EnvironmentalConditions.find('Humidity')
        AirDensity = EnvironmentalConditions.find('AirDensity')
        Methods = WeightSetCalibration.find('Methods')
        Method = Methods.findall('Method')
        # Признак метода Калибровка или Поверка
        Method_ID = Method[0].text
        ReferenceWeightSets = WeightSetCalibration.find('ReferenceWeightSets')
        ReferenceWeightSet = ReferenceWeightSets.findall('ReferenceWeightSet')

        ReferenceWeights = WeightSetCalibration.find('ReferenceWeights')
        ReferenceWeight = ReferenceWeights.findall('ReferenceWeight')

        MassComparators = WeightSetCalibration.find('MassComparators')
        MassComparator = MassComparators.findall('MassComparator')

        TestWeightCalibrations_Count = TestWeightCalibrations.get('Count')  # количество гирь
        EndDate = WeightSetCalibration.get('EndDate')  # дата поверки
        # CertificateNumber = WeightSetCalibration.get('CertificateNumber')  # номер сертификата
        CalibratedBy = WeightSetCalibration.get('CalibratedBy')  # поверитель
        CustomerNumber = ContactOwner.get('CustomerNumber')  # ИНН
        if Department == 0:
            Company_Name = Company  # назвение заказчика
        else:
            Company_Name = Company + ' ' + Department

        # Address = City.get('ZipCode') + ', ' + City.get('State') + ', ' + ContactOwner.find('Address').text  # адрес
        AirDensity_Min = self.roundStr(AirDensity.find('Min').text, 4)
        AirDensity_Max = self.roundStr(AirDensity.find('Max').text, 4)
        AirDensity_Avr = self.roundStr(AirDensity.find('Average').text, 4)

        AirTemperature_Min = self.roundStr(AirTemperature.get('Min'), 2)  # температура мин
        AirTemperature_Max = self.roundStr(AirTemperature.get('Max'), 2)  # температура макс
        AirTemperature_Avr = self.roundStr(AirTemperature.get('Average'), 2)  # температура средняя
        # AirTemperature_Unit = AirTemperature.get('Unit')  # размерность температуры

        AirPressure_Min = self.roundStr(AirPressure.get('Min'), 2)  # давление мин
        AirPressure_Max = self.roundStr(AirPressure.get('Max'), 2)  # давление макс
        AirPressure_Avr = self.roundStr(AirPressure.get('Average'), 2)  # давление среднее
        # AirPressure_Unit = AirPressure.get('Unit')  # размерность давления

        Humidity_Min = self.roundStr(Humidity.get('Min'), 2)  # влажность мин
        Humidity_Max = self.roundStr(Humidity.get('Max'), 2)  # влажность макс
        Humidity_Avr = self.roundStr(Humidity.get('Average'), 2)  # влажность средняя

        Humidity_Unit = Humidity.get('Unit')  # размерность влажности

        Method_Name = Method[0].get('Name')  # метод поверки
        Method_Process = Method[0].get('Process')  # название процесса поверки

        MassComparator_Model = MassComparator[0].get('Model')  # модель компаратора
        MassComparator_SerialNumber = MassComparator[0].get('SerialNumber')  # серийный номер компаратора
        MassComparator_Description = MassComparator[0].find(
            'Description').text  # описание компаратора (дискретность, ско...)

        # ReferenceWeightSet_SerialNumber = ReferenceWeightSet[0].get('SerialNumber')  # серийный номер набора эталонов
        Numbers = str(ReferenceWeightSet[0].get('SerialNumber')).split(' ')
        ReferenceWeightSet_SerialNumber = Numbers[0]
        if len(Numbers) > 1:
            ReferenceWeightSet_RegNumber = Numbers[1]
        else:
            ReferenceWeightSet_RegNumber = ""
        ReferenceWeightSet_Class = ReferenceWeightSet[0].get('Class')  # класс набора эталонов
        ReferenceWeightSet_Range = ReferenceWeightSet[0].get('Range')  # диапазон набора эталонов
        # дата следующей калибровки эталонов
        ReferenceWeightSet_NextCalibrationDate = ReferenceWeightSet[0].get('NextCalibrationDate')

        # ReferenceWeight_Array = []  # массив наборов эталонов
        Test_Passed = True

        # Есть отрицательные результаты или ошибочно записанные AsFound
        if len(TestWeightCalibrationAsFound) > 0:
            for found in TestWeightCalibrationAsFound:
                Nominal = float(str(found.find('Nominal').text).replace(',', '.'))
                Error = float(str(found.find('ConventionalMassCorrection').text).replace(',', '.'))
                Tolerance = float(str(found.find('Tolerance').text).replace(',', '.'))
                # Отрицательный результат
                if abs(Error) < 0.1 * Nominal * 1000 and abs(Error) > Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = False
                # Ошибочно записанный положительный результат
                elif abs(Error) <= Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = True

        # Заполняем протокол поверки если есть положительные результаты AsReturned
        if len(TestWeightCalibrationAsReturned) > 0:

            rb = xlrd.open_workbook(self.TemplateApprovalProtocol, formatting_info=True,
                                    on_demand=True)  # открываем книгу
            wb = xlcopy(rb)  # копируем книгу в память
            ws = wb.get_sheet(0)  # выбираем лист протокола поверки

            rbCert = xlrd.open_workbook(self.TemplateApprovalCert, formatting_info=True,
                                        on_demand=True)  # открываем книгу
            wbCert = xlcopy(rbCert)  # копируем книгу в память
            wsCert = wbCert.get_sheet(0)  # выбираем лист свидетельства
            wsCertMert = wbCert.get_sheet(0)  # выбираем лист метрологических характеристик

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
            styleCellBottom = xlwt.easyxf('border: bottom thin')

            styleCellBorder = xlwt.easyxf('border: left thin, right thin')

            styleCellLeftBottom = xlwt.easyxf('border: left thin, bottom thin, right thin; align: horiz left')

            if TestWeightCalibrations_Count == '1':
                CI_Name = 'Гиря'
            else:
                CI_Name = 'Набор гирь'

            # КрасЦСМ номер протокола не печатаем
            if self.CSM != "КрасЦСМ":
                ws.write(1, 4, self.ApprovalProtocolNum)  # номер протокола
                self.setNums(str(int(self.ApprovalProtocolNum) + 1), 'ApprovalProtocolNum')
            ws.write(2, 1, EndDate)  # дата поверки
            ws.write(3, 1, CI_Name)  # наименование СИ
            ws.write(2, 6, TestWeightSet_AccuracyClass)  # класс точности
            ws.write(3, 6, TestWeightSet_Range)  # номинальное заначение массы
            ws.write(4, 6, TestWeightSet_SerialNumber)  # серийный номер
            ws.write(4, 1, TestWeightSet_Description)  # год выпуска
            ws.write(5, 1, Company_Name)  # название заказчика
            ws.write(6, 1, CustomerNumber)  # номер заказчика
            ws.write(7, 1, TestWeightSet_Manufacturer)  # производитель гирь
            ws.write(9, 1, TestWeightSet_InternalID)  # серия и номер клейма предыдущей поверки

            ws.write(14, 2, Density)  # плотность материала гирь

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
            # TODO: Метрологические характеристики набора
            for ref in ReferenceWeightSet:
                # название набора гирь
                if ref.get('Range') != "":
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    # ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, ReferenceWeightSet_SerialNumber, styleCellCenter)
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, ReferenceWeightSet_SerialNumber, styleCellCenter)
                row += 1

            for comp in MassComparator:
                # название компаратора
                MassComparator_Model = comp.get('Model')
                ws.write(row, 0, MassComparator_Model, styleCellCenter)
                MassComparator_SerialNumber = comp.get('SerialNumber')
                ws.write(row, 2, MassComparator_SerialNumber, styleCellCenter)
                MassComparator_Description = comp.find('Description').text
                # описание компаратора. В поле Описание (Description) должны быть записаны дискретность и СКО модели компаратора
                ws.write(row, 4, MassComparator_Description, styleCellCenter)
                row += 1
            # TODO: Настройки в шаблон
            Row = 46
            RowMetr = 5
            for i in TestWeightCalibrationAsReturned:
                Nominal = i.find('Nominal').text
                NominalUnit = i.find('NominalUnit').text
                if NominalUnit == 'g':
                    NominalUnit = 'г'
                if NominalUnit == 'mg':
                    NominalUnit = 'мг'
                if NominalUnit == 'kg':
                    NominalUnit = 'кг'

                WeightID = i.find('WeightID').text
                Index = i.find('ReferenceWeight').text
                ReferenceWeight_ConventionalMassError = 0
                Tolerance = i.find('Tolerance').text
                for j in ReferenceWeight:
                    if Index == j.get('Index'):
                        ReferenceWeight_ConventionalMassError = j.get('ConventionalMassError')
                if TestWeightCalibrations_Count == '1':
                    ws.write(Row, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
                    # wsCert.write(RowMetr, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
                else:
                    ws.write(Row, 0, str.strip(WeightID + Nominal + NominalUnit), styleCellCenterSinglCellTop)
                    # wsCert.write(RowMetr, 0, str.strip(WeightID + Nominal + NominalUnit), styleCellCenterSinglCellTop)
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
                        A1.append(WeightReading[cicle * 3].get('WeightReading'))
                        WeightReadingUnit = WeightReading[cicle].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle * 3 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle * 3 + 2].get('WeightReading'))
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
                        A1.append(WeightReading[cicle * 4].get('WeightReading'))
                        WeightReadingUnit = WeightReading[0].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle * 4 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B2.append(WeightReading[cicle * 4 + 2].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B2 ' + B2[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle * 4 + 3].get('WeightReading'))
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
                ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection, 4)
                ConventionalMassCorrectionUnit = i.find('ConventionalMassCorrectionUnit').text
                ConventionalMass = i.find('ConventionalMass').text
                # ConventionalMass = self.roundStr(ConventionalMass,4)
                ConventionalMassUnit = i.find('ConventionalMassUnit').text

                if ConventionalMassUnit == 'g':
                    ConventionalMassUnit = 'г'
                if ConventionalMassUnit == 'mg':
                    ConventionalMassUnit = 'мг'
                if ConventionalMassUnit == 'kg':
                    ConventionalMassUnit = 'кг'

                if ConventionalMassCorrectionUnit == 'g':
                    ConventionalMassCorrectionUnit = 'г'
                if ConventionalMassCorrectionUnit == 'mg':
                    ConventionalMassCorrectionUnit = 'мг'
                if ConventionalMassCorrectionUnit == 'kg':
                    ConventionalMassCorrectionUnit = 'кг'

                ExpandedMassUncertainty = i.find('ExpandedMassUncertainty').text
                ExpandedMassUncertaintyUnit = i.find('ExpandedMassUncertaintyUnit').text
                ws.write(Row, 3, Avr, styleCellCenterSinglCellTop)
                ReferenceWeight_ConventionalMassError = self.roundStr(ReferenceWeight_ConventionalMassError, 4)
                ws.write(Row, 4, ReferenceWeight_ConventionalMassError,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 5, ConventionalMassCorrection,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 6, ConventionalMass + ConventionalMassUnit, styleCellCenterSinglCellTop)

                wsCert.write(RowMetr, 11, WeightID + Nominal + NominalUnit, styleCellCenter)
                wsCert.write(RowMetr, 12, "", styleCellCenter)
                wsCert.write(RowMetr, 13, ConventionalMass + ConventionalMassUnit, styleCellCenter)
                wsCert.write(RowMetr, 14, "", styleCellCenter)
                wsCert.write(RowMetr, 15, ConventionalMassCorrection + ConventionalMassCorrectionUnit, styleCellCenter)
                wsCert.write(RowMetr, 16, "", styleCellCenter)
                wsCert.write(RowMetr, 17, ExpandedMassUncertainty, styleCellCenter)
                wsCert.write(RowMetr, 18, "", styleCellLeftSinglCell)

                ws.write(Row, 7, ExpandedMassUncertainty, styleCellCenterSinglCellTop)

                for x in range(Row + 1, Row + len(WeightReading)):
                    ws.write(x, 0, '', styleCellLeftLine)
                    ws.write(x, 8, '', styleCellRightLine)
                Row = RowWeightReading
                RowMetr += 1
            for y in range(0, 9):
                ws.write(Row, y, '', styleCellTopLine)

            if Test_Passed:
                ws.write(Row + 2, 0,
                         'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                ws.write(Row + 4, 0, 'На основании результатов поверки выдано свидетельство о поверке № ' + str(
                    self.ApprovalCertNum) + ' от _____._____________._______г.')
            else:
                if self.CSM != 'КрасЦСМ':
                    ws.write(Row + 2, 0,
                             'Заключение по результатам поверки: гири не пригодны к использованию по классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 4, 0, 'На основании результатов поверки выдано извещение о непригодности № ' + str(
                        self.ErrorNum) + ' от _____._____________._______г.')
                else:
                    ws.write(Row + 2, 0,
                             'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 4, 0,
                             'На основании результатов поверки выдано свидетельство о поверке № _______________________ от _____._____________._______г.')

                pass

            ws.write(Row + 7, 0, 'Поверитель:_____________________ ' + CalibratedBy)
            ws.write(Row + 7, 6, 'Дата протокола: ' + str(datetime.today().strftime("%d.%m.%Y")))

            ws.insert_bitmap('logo.bmp', 1, 7)

            # сохранение данных в новый документ
            date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
            file_to_save = self.rightFileName(
                TestWeightSet_Comment.strip() + ' ' + Company_Name.strip() + ' ' + TestWeightSet_AccuracyClass.strip() + ' ' + TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.xls')

            # Составление свидетельства о поверке

            fileApprovalProtocol = self.Excel_folder + '\\' + self.ApprovalProtocolFolder + '\\' + file_to_save
            fileApprovalProtocol = fileApprovalProtocol.replace(',', ' ')

            if Test_Passed == True:
                NextYear = datetime.strptime(str(EndDate), "%d.%m.%Y").date()
                NextYear += relativedelta(years=1)
                wsCert.write(2, 4, self.ApprovalCertNum, styleCellBottom)
                self.setNums(str(int(self.ApprovalCertNum) + 1), 'ApprovalCertNum')
                wsCert.write(4, 8, str(NextYear.strftime("%d.%m.%Y")), styleCellBottom)
                wsCert.write(6, 2,
                             CI_Name + " " + TestWeightSet_Range + " " + TestWeightSet_AccuracyClass + " " + TestWeightSet_Comment,
                             styleCellBottom)
                wsCert.write(12, 4, TestWeightSet_InternalID,
                             styleCellBottom)  # серия и номер клейма предыдущей поверки
                # wsCert.write(10, 2, TestWeightSet_SerialNumber, styleCellBottom)
                wsCert.write(14, 2, TestWeightSet_SerialNumber, styleCellBottom)
                # wsCert.write(14, 1, Company_Name + ", ИНН " + CustomerNumber, styleCellBottom)
                Method_ID = str(Method_ID).strip('Поверка')
                wsCert.write(16, 2, Method_ID, styleCellBottom)
                if ReferenceWeightSet_RegNumber != "":
                    ReferenceWeightSet_RegNumber = ", рег.№ " + ReferenceWeightSet_RegNumber
                wsCert.write(20, 2,
                             ReferenceWeightSet_info + ", заводской номер " + ReferenceWeightSet_SerialNumber + ReferenceWeightSet_RegNumber)
                wsCert.write(28, 0, AirTemperature_Avr + " oC,  отностительная влажность " + Humidity_Avr + "%",
                             styleCellBottom)

                fileApprovalCert = self.Excel_folder + '\\' + self.ApprovalCertFolder + '\\Свидетельство о поверке ' + file_to_save
                fileApprovalCert = fileApprovalCert.replace(',', ' ')

                if self.ApprovalCertEnable == False:
                    wb.save(fileApprovalProtocol)
                    if self.autoopen == True:
                        os.startfile(fileApprovalProtocol, 'open')

                else:
                    wbCert.save(fileApprovalCert)
                    wb.save(fileApprovalProtocol)
                    if self.autoopen == True:
                        os.startfile(fileApprovalCert)
                        sleep(1)
                        os.startfile(fileApprovalProtocol)

            else:  # составляем извещение о непригодности
                rbError = xlrd.open_workbook(self.TemplateError, formatting_info=True,
                                             on_demand=True)  # открываем книгу
                wbError = xlcopy(rbError)  # копируем книгу в память
                wsError = wbError.get_sheet(0)  # выбираем лист извещения о непригодности
                wsError.write(3, 4, self.ErrorNum, styleCellBottom)
                self.setNums(str(int(self.ErrorNum) + 1), 'ErrorNum')
                wsError.write(7, 2,
                              CI_Name + " " + TestWeightSet_Range + " " + TestWeightSet_AccuracyClass + " " + TestWeightSet_Comment,
                              styleCellBottom)
                wsError.write(13, 4, TestWeightSet_InternalID, styleCellBottom)
                wsError.write(15, 2, TestWeightSet_SerialNumber, styleCellBottom)
                wsError.write(17, 2, Company_Name + ", ИНН " + CustomerNumber, styleCellBottom)
                wsError.write(34, 0, str(datetime.today().strftime("%d.%m.%Y")), styleCellBottom)

                fileError = self.Excel_folder + '\\' + self.ErrorFolder + '\\Извещение о непригодности ' + file_to_save
                if self.ErrorEnable:
                    wbError.save(fileError)
                    if self.autoopen == True:
                        os.startfile(fileError)
                if self.CalProtocolEnable:
                    wb.save(fileApprovalProtocol)
                    if self.autoopen == True:
                        os.startfile(fileApprovalProtocol)

    # формирование протокола калибровки
    def CallReport(self, xml_filename):
        tree = etree.parse(xml_filename)
        root = tree.getroot()

        WeightSetCalibration = root.find('WeightSetCalibration')
        ContactOwner = WeightSetCalibration.find('ContactOwner')
        Company = str(ContactOwner.find('Company').text).strip(' ')
        Department = str(ContactOwner.find('Department').text).strip(' ')

        City = ContactOwner.find('City')

        TestWeightSet = WeightSetCalibration.find('TestWeightSet')
        TestWeightSet_Description = TestWeightSet.find('Description').text  # год выпуста
        TestWeightSet_Comment = TestWeightSet.find('Comment').text  # номер в ГР
        TestWeightSet_SerialNumber = TestWeightSet.get('SerialNumber')
        TestWeightSet_AccuracyClass = TestWeightSet.get('AccuracyClass')
        TestWeightSet_Manufacturer = TestWeightSet.get('Manufacturer')
        TestWeightSet_InternalID = TestWeightSet.get('InternalID')  # номер клейма предыдущей поверки
        TestWeightSet_Range = TestWeightSet.get('Range')

        str(TestWeightSet_Range).replace('g', 'г')
        str(TestWeightSet_Range).replace('kg', 'кг')
        str(TestWeightSet_Range).replace('mg', 'мг')

        TestWeightCalibrations = TestWeightSet.find('TestWeightCalibrations')
        TestWeightCalibrationAsReturned = TestWeightCalibrations.findall('TestWeightCalibrationAsReturned')
        TestWeightCalibrationAsFound = TestWeightCalibrations.findall('TestWeightCalibrationAsFound')
        TestWeightSet_AlloyMaterials = TestWeightSet.find('AlloyMaterials')
        TestWeightSet_AlloyMaterial = TestWeightSet_AlloyMaterials.findall('AlloyMaterial')[0]
        Density = TestWeightSet_AlloyMaterial.get('Density') + TestWeightSet_AlloyMaterial.get('DensityUnit')

        EnvironmentalConditions = WeightSetCalibration.find('EnvironmentalConditions')
        AirTemperature = EnvironmentalConditions.find('AirTemperature')
        AirPressure = EnvironmentalConditions.find('AirPressure')
        Humidity = EnvironmentalConditions.find('Humidity')
        AirDensity = EnvironmentalConditions.find('AirDensity')
        Methods = WeightSetCalibration.find('Methods')
        Method = Methods.findall('Method')
        # Признак метода Калибровка или Поверка
        Method_ID = Method[0].text
        CallReport = False
        if str(Method_ID).find('Калибровка') != -1:
            CallReport = True

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
        AirDensity_Min = self.roundStr(AirDensity.find('Min').text, 4)
        AirDensity_Max = self.roundStr(AirDensity.find('Max').text, 4)
        AirDensity_Avr = self.roundStr(AirDensity.find('Average').text, 4)

        AirTemperature_Min = self.roundStr(AirTemperature.get('Min'), 2)  # температура мин
        AirTemperature_Max = self.roundStr(AirTemperature.get('Max'), 2)  # температура макс
        AirTemperature_Avr = self.roundStr(AirTemperature.get('Average'), 2)  # температура средняя
        AirTemperature_Unit = AirTemperature.get('Unit')  # размерность температуры

        AirPressure_Min = self.roundStr(AirPressure.get('Min'), 2)  # давление мин
        AirPressure_Max = self.roundStr(AirPressure.get('Max'), 2)  # давление макс
        AirPressure_Avr = self.roundStr(AirPressure.get('Average'), 2)  # давление среднее
        AirPressure_Unit = AirPressure.get('Unit')  # размерность давления

        Humidity_Min = self.roundStr(Humidity.get('Min'), 2)  # влажность мин
        Humidity_Max = self.roundStr(Humidity.get('Max'), 2)  # влажность макс
        Humidity_Avr = self.roundStr(Humidity.get('Average'), 2)  # влажность средняя

        Humidity_Unit = Humidity.get('Unit')  # размерность влажности

        Method_Name = Method[0].get('Name')  # метод поверки
        Method_Process = Method[0].get('Process')  # название процесса поверки

        MassComparator_Model = MassComparator[0].get('Model')  # модель компаратора
        MassComparator_SerialNumber = MassComparator[0].get('SerialNumber')  # серийный номер компаратора
        MassComparator_Description = MassComparator[0].find(
            'Description').text  # описание компаратора (дискретность, ско...)

        # ReferenceWeightSet_SerialNumber = ReferenceWeightSet[0].get('SerialNumber')  # серийный номер набора эталонов
        Numbers = str(ReferenceWeightSet[0].get('SerialNumber')).split(' ')
        ReferenceWeightSet_SerialNumber = Numbers[0]
        if len(Numbers) > 1:
            ReferenceWeightSet_RegNumber = Numbers[1]
        else:
            ReferenceWeightSet_RegNumber = ""
        ReferenceWeightSet_Class = ReferenceWeightSet[0].get('Class')  # класс набора эталонов
        ReferenceWeightSet_Range = ReferenceWeightSet[0].get('Range')  # диапазон набора эталонов
        ReferenceWeightSet_NextCalibrationDate = ReferenceWeightSet[0].get(
            'NextCalibrationDate')  # дата следующей калибровки эталонов
        ReferenceWeightSet_CertificateNumber = ReferenceWeightSet[0].get(
            'CertificateNumber')  # ReferenceWeight_Array = []  # массив наборов эталонов
        ReferenceWeightSet_Comment = ReferenceWeightSet[0].find('Comment').text
        Test_Passed = True

        # Есть отрицательные результаты или ошибочно записанные AsFound
        if len(TestWeightCalibrationAsFound) > 0:
            for found in TestWeightCalibrationAsFound:
                Nominal = float(str(found.find('Nominal').text).replace(',', '.'))
                Error = float(str(found.find('ConventionalMassCorrection').text).replace(',', '.'))
                Tolerance = float(str(found.find('Tolerance').text).replace(',', '.'))
                # Отрицательный результат
                if abs(Error) < 0.1 * Nominal * 1000 and abs(Error) > Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = False
                # Ошибочно записанный положительный результат
                elif abs(Error) <= Tolerance:
                    TestWeightCalibrationAsReturned.append(found)
                    Test_Passed = True

        # Есть положительные результаты AsReturned
        if len(TestWeightCalibrationAsReturned) > 0:

            rb = xlrd.open_workbook(self.TemplateCalProtocol, formatting_info=True, on_demand=True)  # открываем книгу
            wb = xlcopy(rb)  # копируем книгу в память

            # Печать протокола калибровки
            ws = wb.get_sheet(0)  # выбираем лист протокола поверки

            rbCert = xlrd.open_workbook(self.TemplateCalCert, formatting_info=True, on_demand=True)  # открываем книгу
            wbCert = xlcopy(rbCert)  # копируем книгу в память

            # Печать протокола калибровки
            wsCert = wbCert.get_sheet(0)  # выбираем лист сертификата
            wsCertMetr = wbCert.get_sheet(0)  # выбираем лист метрологических характеристик

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
            styleCellBottom = xlwt.easyxf('border: bottom thin')
            styleCellBorder = xlwt.easyxf('border: left thin, right thin')
            styleCellBottomCenter = xlwt.easyxf('border: bottom thin; align: horiz center')
            styleCellLeftBottom = xlwt.easyxf('border: left thin, bottom thin, right thin; align: horiz left')

            if TestWeightCalibrations_Count == '1':
                CI_Name = 'Гиря'
            else:
                CI_Name = 'Набор гирь'

            # КрасЦСМ номер протокола не печатаем
            if self.CSM != "КрасЦСМ":
                ws.write(1, 4, self.CalProtocolNum, styleCellBottomCenter)  # номер протокола
                self.setNums(str(int(self.CalProtocolNum) + 1), 'CalProtocolNum')
            ws.write(2, 1, EndDate)  # дата поверки

            ws.write(3, 1, CI_Name)  # наименование СИ

            ws.write(2, 6, TestWeightSet_AccuracyClass)  # класс точности
            ws.write(3, 6, TestWeightSet_Range)  # номинальное заначение массы
            ws.write(4, 6, TestWeightSet_SerialNumber)  # серийный номер

            ws.write(4, 1, TestWeightSet_Description)  # год выпуска
            ws.write(5, 1, Company_Name)  # название заказчика
            ws.write(6, 1, CustomerNumber)  # номер заказчика
            ws.write(7, 1, TestWeightSet_Manufacturer)  # производитель гирь

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

            # TODO: Настройки в шаблон
            row = 37
            # TODO: Метрологические характеристики набора
            for ref in ReferenceWeightSet:
                # название набора гирь
                if ref.get('Range') != "":
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    # ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, ReferenceWeightSet_SerialNumber, styleCellCenter)
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, ReferenceWeightSet_SerialNumber, styleCellCenter)
                row += 1

            for comp in MassComparator:
                # название компаратора
                MassComparator_Model = comp.get('Model')
                ws.write(row, 0, MassComparator_Model, styleCellCenter)
                MassComparator_SerialNumber = comp.get('SerialNumber')
                ws.write(row, 2, MassComparator_SerialNumber, styleCellCenter)
                MassComparator_Description = comp.find('Description').text
                # описание компаратора. В поле Описание (Description) должны быть записаны дискретность и СКО модели компаратора
                ws.write(row, 4, MassComparator_Description, styleCellCenter)
                row += 1
            # TODO: Настройки в шаблон
            Row = 46
            RowMetr = 5
            for i in TestWeightCalibrationAsReturned:
                Nominal = i.find('Nominal').text
                NominalUnit = i.find('NominalUnit').text
                if NominalUnit == 'g':
                    NominalUnit = 'г'
                if NominalUnit == 'mg':
                    NominalUnit = 'мг'
                if NominalUnit == 'kg':
                    NominalUnit = 'кг'

                WeightID = i.find('WeightID').text
                Index = i.find('ReferenceWeight').text
                ReferenceWeight_ConventionalMassError = 0
                Tolerance = i.find('Tolerance').text
                for j in ReferenceWeight:
                    if Index == j.get('Index'):
                        ReferenceWeight_ConventionalMassError = j.get('ConventionalMassError')
                if TestWeightCalibrations_Count == '1':
                    ws.write(Row, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
                    # wsCertMetr.write(RowMetr, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
                else:
                    ws.write(Row, 0, str.strip(WeightID + Nominal + NominalUnit), styleCellCenterSinglCellTop)
                    # wsCertMetr.write(RowMetr, 0, str.strip(WeightID + Nominal + NominalUnit), styleCellCenterSinglCellTop)
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
                        A1.append(WeightReading[cicle * 3].get('WeightReading'))
                        WeightReadingUnit = WeightReading[cicle].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle * 3 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle * 3 + 2].get('WeightReading'))
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
                        A1.append(WeightReading[cicle * 4].get('WeightReading'))
                        WeightReadingUnit = WeightReading[0].get('WeightReadingUnit')
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + A1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B1.append(WeightReading[cicle * 4 + 1].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + B1[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        B2.append(WeightReading[cicle * 4 + 2].get('WeightReading'))
                        ws.write(RowWeightReading, 1, str(cicle + 1) + ' B2 ' + B2[cicle],
                                 styleCellLeftSinglCell)
                        RowWeightReading += 1
                        A2.append(WeightReading[cicle * 4 + 3].get('WeightReading'))
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
                ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection, 4)
                ConventionalMassCorrectionUnit = i.find('ConventionalMassCorrectionUnit').text
                ConventionalMass = i.find('ConventionalMass').text
                # ConventionalMass = self.roundStr(ConventionalMass,4)
                ConventionalMassUnit = i.find('ConventionalMassUnit').text

                if ConventionalMassUnit == 'g':
                    ConventionalMassUnit = 'г'
                if ConventionalMassUnit == 'mg':
                    ConventionalMassUnit = 'мг'
                if ConventionalMassUnit == 'kg':
                    ConventionalMassUnit = 'кг'

                if ConventionalMassCorrectionUnit == 'g':
                    ConventionalMassCorrectionUnit = 'г'
                if ConventionalMassCorrectionUnit == 'mg':
                    ConventionalMassCorrectionUnit = 'мг'
                if ConventionalMassCorrectionUnit == 'kg':
                    ConventionalMassCorrectionUnit = 'кг'

                ExpandedMassUncertainty = i.find('ExpandedMassUncertainty').text
                # ExpandedMassUncertaintyUnit = i.find('ExpandedMassUncertaintyUnit').text
                ws.write(Row, 3, Avr, styleCellCenterSinglCellTop)
                ReferenceWeight_ConventionalMassError = self.roundStr(ReferenceWeight_ConventionalMassError, 4)
                ws.write(Row, 4, ReferenceWeight_ConventionalMassError,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 5, ConventionalMassCorrection,
                         styleCellCenterSinglCellTop)
                ws.write(Row, 6, ConventionalMass + ConventionalMassUnit, styleCellCenterSinglCellTop)
                ws.write(Row, 7, ExpandedMassUncertainty, styleCellCenterSinglCellTop)

                wsCert.write(RowMetr, 11, WeightID + Nominal + NominalUnit, styleCellCenter)
                wsCert.write(RowMetr, 12, "", styleCellCenter)
                wsCert.write(RowMetr, 13, ConventionalMass + ConventionalMassUnit, styleCellCenter)
                wsCert.write(RowMetr, 14, "", styleCellCenter)
                wsCert.write(RowMetr, 15, ConventionalMassCorrection + ConventionalMassCorrectionUnit, styleCellCenter)
                wsCert.write(RowMetr, 16, "", styleCellCenter)
                wsCert.write(RowMetr, 17, ExpandedMassUncertainty, styleCellCenter)
                wsCert.write(RowMetr, 18, "", styleCellLeftSinglCell)
                RowMetr += 1
                for x in range(Row + 1, Row + len(WeightReading)):
                    ws.write(x, 0, '', styleCellLeftLine)
                    ws.write(x, 8, '', styleCellRightLine)
                Row = RowWeightReading

            for y in range(0, 9):
                ws.write(Row, y, '', styleCellTopLine)

            if Test_Passed:
                ws.write(Row + 2, 0,
                         'Заключение по результатам калибровки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                ws.write(Row + 4, 0, 'На основании результатов калибровки выдан сертификат о калибровке № ' + str(
                    self.CalCertNum) + '        от  _____._____________._______г.')
            else:
                if self.CSM != 'КрасЦСМ':
                    ws.write(Row + 2, 0,
                             'Заключение по результатам калибровки: гири не пригодны к использованию по классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 4, 0,
                             'На основании результатов поверки выдано извещение о непригодности № _______________________ от _____._____________._______г.')
                else:
                    ws.write(Row + 2, 0,
                             'Заключение по результатам калибровки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 4, 0,
                             'На основании результатов калибровки выдан сертификат о калибровке № ______________________  от _____._____________._______г.')

                pass

            ws.write(Row + 7, 0, 'Поверитель:_____________________ ' + CalibratedBy)
            ws.write(Row + 7, 6, 'Дата протокола: ' + str(datetime.today().strftime("%d.%m.%Y")))

            ws.insert_bitmap('logo.bmp', 1, 7)

            # Составление сертификата калибровки

            NextYear = datetime.strptime(str(EndDate), "%d.%m.%Y").date()
            NextYear += relativedelta(years=1)
            wsCert.write(2, 4, str(self.CalCertNum), styleCellBottomCenter)
            self.setNums(str(int(self.CalCertNum) + 1), 'CalCertNum')
            wsCert.write(4, 8, str(NextYear.strftime("%d.%m.%Y")), styleCellBottom)
            wsCert.write(6, 2, CI_Name + " " + TestWeightSet_Range + " " + TestWeightSet_Comment, styleCellBottom)
            wsCert.write(10, 2, TestWeightSet_SerialNumber, styleCellBottom)
            wsCert.write(12, 1, TestWeightSet_Manufacturer, styleCellBottom)
            wsCert.write(14, 1, Company_Name + ", ИНН " + CustomerNumber, styleCellBottom)
            Method_ID = str(Method_ID).strip('Калибровка')
            wsCert.write(20, 2, Method_ID, styleCellBottom)
            wsCert.write(28, 0, AirTemperature_Avr + " oC,  отностительная влажность " + Humidity_Avr + "%",
                         styleCellBottom)

            if ReferenceWeightSet_RegNumber != "":
                ReferenceWeightSet_RegNumber = ", рег. № " + ReferenceWeightSet_RegNumber

            wsCert.write(30, 2,
                         ReferenceWeightSet_info + ", заводской номер " + ReferenceWeightSet_SerialNumber + ReferenceWeightSet_RegNumber,
                         styleCellBottom)

            # Сохранение сертификата и протокола
            # сохранение данных в новый документ
            date_time = datetime.now().strftime("%d%m%Y_%H%M%S")

            file_to_save = self.rightFileName(
                TestWeightSet_Comment.strip() + ' ' + Company_Name.strip() + ' ' + TestWeightSet_AccuracyClass.strip() + ' ' + TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.xls')
            fileCalProtocol = self.Excel_folder + '\\' + self.CalProtocolFolder + '\\' + file_to_save
            fileCalCert = self.Excel_folder + '\\' + self.CalCertFolder + '\\Сертификат о калибровке ' + file_to_save

            if self.CalCertEnable == False:
                wb.save(fileCalProtocol)
                if self.autoopen == True:
                    os.startfile(fileCalProtocol)
            else:
                wb.save(fileCalProtocol)
                wbCert.save(fileCalCert)
                if self.autoopen == True:
                    os.startfile(fileCalProtocol)
                    os.startfile(fileCalCert)

    # запуск конвертации
    def convertation(self, xml_filename=None):

        tree = etree.parse(xml_filename)
        root = tree.getroot()
        WeightSetCalibration = root.find('WeightSetCalibration')
        Methods = WeightSetCalibration.find('Methods')
        Method = Methods.findall('Method')
        # Признак метода Калибровка или Поверка
        Method_ID = Method[0].text

        IsCallReport = False
        IsApprovalReport = False

        if str(Method_ID).find('Калибровка') != -1:
            IsCallReport = True
            IsApprovalReport = False
        else:
            IsCallReport = False
            IsApprovalReport = True

        if IsApprovalReport == True:
            self.ApprovalReport(xml_filename)
        if IsCallReport == True:
            self.CallReport(xml_filename)
        # Удаляем исходный файл xml
        # if self.autodelXML == True:
        os.remove(xml_filename)


# класс главного окна
class MainWindow(QMainWindow):
    demon = DemonConvertation()
    Title = " Сохранение протоколов поверки MCLink v" + ver

    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.demon.update_settings()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.btScanFolder.clicked.connect(self.selectXmlFolder)
        self.ui.btDestFolder.clicked.connect(self.selectExcelFolder)

        self.ui.btTemplateApprovalProtocol.clicked.connect(self.selectTemplateApprovalProtocol)
        self.ui.btTemplateApprovalCert.clicked.connect(self.selectTemplateApprovalCert)
        self.ui.btTemplateCalProtocol.clicked.connect(self.selectTemplateCalProtocol)
        self.ui.btTemplateCalCert.clicked.connect(self.selectTemplateCalCert)
        self.ui.btTemplateError.clicked.connect(self.selectTemplateError)

        self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalProtocol)
        self.ui.lbTemplateApprovalCert.setText(self.demon.TemplateApprovalCert)
        self.ui.lbTemplateCalProtocol.setText(self.demon.TemplateCalProtocol)
        self.ui.lbTemplateCalCert.setText(self.demon.TemplateCalCert)
        self.ui.lbTemplateError.setText(self.demon.TemplateError)

        self.ui.lbScanFolder.setText(self.demon.xml_folder)
        self.ui.lbDestFolder.setText(self.demon.Excel_folder)
        self.ui.lbTemplateApprovalProtocol.setText(self.demon.template_filename)
        self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalProtocol)
        self.ui.lbTemplateApprovalCert.setText(self.demon.TemplateApprovalCert)
        self.setWindowTitle(self.demon.CSM + self.Title)

        self.ui.edTemplateReserved.setVisible(False)
        self.ui.exitAction.triggered.connect(self.close)
        self.ui.startAction.triggered.connect(self.start)
        self.ui.startAction.setVisible(True)
        self.ui.stopAction.triggered.connect(self.stop)
        self.ui.stopAction.setVisible(False)

        self.ui.chbAutoOpen.setChecked(self.demon.autoopen)
        self.ui.chbAutoOpen.clicked.connect(self.changeAutoOpen)
        self.ui.chbAutoDelXML.setVisible(False)
        # self.statusBar()
        self.ui.edApprovalProtocolNum.setText(str(self.demon.ApprovalProtocolNum))
        self.ui.edApprovalCertNum.setText(str(self.demon.ApprovalCertNum))
        self.ui.edCalProtocolNum.setText(str(self.demon.CalProtocolNum))
        self.ui.edCalCertNum.setText(str(self.demon.CalCertNum))
        self.ui.edErrorNum.setText(str(self.demon.ErrorNum))

        self.ui.btSaveSettings.clicked.connect(self.saveSettings)
        self.ui.cbxApprovalProtocol.setChecked(self.demon.ApprovalProtocolEnable)
        self.ui.cbxApprovalSert.setChecked(bool(self.demon.ApprovalCertEnable))
        self.ui.cbxError.setChecked(bool(self.demon.ErrorEnable))
        self.ui.cbxCalProtocol.setChecked(bool(self.demon.CalProtocolEnable))
        self.ui.cbxCalSert.setChecked(bool(self.demon.CalCertEnable))

        self.ui.cbxApprovalProtocol.clicked.connect(self.setProtocol)
        self.ui.cbxApprovalSert.clicked.connect(self.setProtocol)
        self.ui.cbxError.clicked.connect(self.setProtocol)
        self.ui.cbxCalProtocol.clicked.connect(self.setProtocol)
        self.ui.cbxCalSert.clicked.connect(self.setProtocol)

        if self.demon.runned == True:
            self.start()

    def setProtocol(self):
        self.demon.setProtocol(1, self.ui.cbxApprovalProtocol.isChecked())
        self.demon.setProtocol(2, self.ui.cbxApprovalSert.isChecked())
        self.demon.setProtocol(3, self.ui.cbxError.isChecked())
        self.demon.setProtocol(4, self.ui.cbxCalProtocol.isChecked())
        self.demon.setProtocol(5, self.ui.cbxCalSert.isChecked())

    def saveSettings(self):
        self.demon.setXmlFolder(self.ui.lbScanFolder.text())
        self.demon.setExcelFolder(self.ui.lbDestFolder.text())
        self.demon.setAutoOpen(self.ui.chbAutoOpen.isChecked())
        self.demon.setAutoDelXML(self.ui.chbAutoDelXML.isChecked())
        self.demon.setAutoStart(self.demon.runned)

        self.demon.setNums(self.ui.edApprovalProtocolNum.text(), 'ApprovalProtocolNum')
        self.demon.setNums(self.ui.edApprovalCertNum.text(), 'ApprovalCertNum')
        self.demon.setNums(self.ui.edCalProtocolNum.text(), 'CalProtocolNum')
        self.demon.setNums(self.ui.edCalCertNum.text(), 'CalCertNum')
        self.demon.setNums(self.ui.edErrorNum.text(), 'ErrorNum')

        self.demon.setTemplateFilename(self.ui.lbTemplateApprovalProtocol.text(), 'TemplateApprovalProtocol')
        self.demon.setTemplateFilename(self.ui.lbTemplateApprovalCert.text(), 'TemplateApprovalCert')
        self.demon.setTemplateFilename(self.ui.lbTemplateCalProtocol.text(), 'TemplateCalProtocol')
        self.demon.setTemplateFilename(self.ui.lbTemplateCalCert.text(), 'TemplateCalCert')
        self.demon.setTemplateFilename(self.ui.lbTemplateError.text(), 'TemplateError')

        self.demon.setProtocol(1, self.ui.cbxApprovalProtocol.isChecked())
        self.demon.setProtocol(2, self.ui.cbxApprovalSert.isChecked())
        self.demon.setProtocol(3, self.ui.cbxError.isChecked())
        self.demon.setProtocol(4, self.ui.cbxCalProtocol.isChecked())
        self.demon.setProtocol(5, self.ui.cbxCalSert.isChecked())

    def start(self):
        self.ui.startAction.setVisible(False)
        self.ui.stopAction.setVisible(True)
        self.demon = DemonConvertation()
        self.demon.setAutoStart(True)
        self.demon.setDaemon(True)
        self.demon.start()

    def stop(self):
        self.ui.startAction.setVisible(True)
        self.ui.stopAction.setVisible(False)
        self.demon.setAutoStart(False)

    def selectXmlFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите исходный файл', '')
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setXmlFolder(folder)
            self.ui.lbScanFolder.setText(self.demon.xml_folder)

    def selectExcelFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите папку сохранения отчетов', '')
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setExcelFolder(folder)
            self.ui.lbDestFolder.setText(self.demon.Excel_folder)

    def selectTemplateApprovalProtocol(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона потокола поверки',
                                                    self.demon.TemplateApprovalProtocol, '*.xls')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateApprovalProtocol')
            self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalProtocol)

    def selectTemplateApprovalCert(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона свидетельства о поверке',
                                                    self.demon.TemplateApprovalCert, '*.xls')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateApprovalCert')
            self.ui.lbTemplateApprovalCert.setText(self.demon.TemplateApprovalCert)

    def selectTemplateError(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона извещения о непригодности',
                                                    self.demon.TemplateError, '*.xls')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateError')
            self.ui.lbTemplateError.setText(self.demon.TemplateError)

    def selectTemplateCalProtocol(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона потокола калибровки', self.demon.TemplateCalProtocol, '*.xls')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateCalProtocol')
            self.ui.lbTemplateCalProtocol.setText(self.demon.TemplateCalProtocol)

    def selectTemplateCalCert(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона сертификата калибровки', self.demon.TemplateCalCert, '*.xls')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateCalCert')
            self.ui.lbTemplateCalCert.setText(self.demon.TemplateCalCert)

    def changeAutoOpen(self):
        if self.ui.chbAutoOpen.isChecked():
            self.demon.setAutoOpen(True)
        else:
            self.demon.setAutoOpen(False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())
