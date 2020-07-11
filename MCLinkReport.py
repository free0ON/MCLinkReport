# -*- coding: utf-8 -*-
"""
Программа конвертации отчетов программы MCLink в формате xml в формат xls
v.3.9.1 Проблема переноса с as returned
v 3.9 Проблема с точкой в массе номинала
        подстрочные символы в имени класса точности
v 3.8 Исправлена ошибка определения набора или гири
v 3.7 Исправлена ошибка в ParseXML
v 3.6 Добавил в лог exception
v 3.5 Добавил лог
v 3.4 Проблемы с отчетом калибровки компаратора
v 3.3 Исправлена проблема с очисткой переменных эталонов
v 3.2 исправлены пути
v 3.1 добавлен лог
v 3.0
v 2.8 Шаблоны свидетельств в формате docx
v 2.7 Список компараторов
v 2.6 Ошибки округления
v 2.5 Добавлена поверка компаратора,
      Устранен баг с остановкой демона при сохранении
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
import decimal
import os
import sys
import xml.etree.ElementTree as etree
from threading import Thread
from time import sleep
import configparser
import xlrd
import xlwt
from statistics import mean, stdev
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from xlutils.copy import copy as xlcopy
from mainwindow import *
from dateutil.relativedelta import relativedelta
from mailmerge import MailMerge
from openpyxl import load_workbook
import openpyxl
import logging
import shutil

ver = "3.9.1"

logfileName = os.path.dirname(sys.argv[0]).replace('/', '\\') + r'\report.log'

logging.basicConfig(filename=logfileName, level=logging.DEBUG, filemode='w', format='%(asctime)s | %(message)s',
                    datefmt='%d.%m.%y %H:%M:%S')


# Подстановка подстрочных символов названий классов
def ClassReName(ClassName):
    if str(ClassName) == "E1":
        return str('\u0045\u2081')
    if str(ClassName) == "F1":
        return str('\u0046\u2081')
    if str(ClassName) == "M1":
        return str('\u004D\u2081')
    if str(ClassName) == "E2":
        return str('\u0045\u2082')
    if str(ClassName) == "F2":
        return str('\u0046\u2082')
    if str(ClassName) == "M2":
        return str('\u004D\u2082')


class TContactOwner:
    def __init__(self, Company, CustomerNumber):
        self.Company = Company
        self.CustomerNumber = CustomerNumber

    Company = ''
    CustomerNumber = ''
    Department = ''
    CustomerName = ''
    EMail = ''
    PhoneNumber = ''
    FaxNumber = ''
    Address = ''


class TEnvironmentalConditions:
    AirTemperature_Min = ''
    AirTemperature_Max = ''
    AirTemperature_Average = ''
    AirTemperature_Unit = ''

    AirPressure_Min = ''
    AirPressure_Max = ''
    AirPressure_Average = ''
    AirPressure_Unit = ''

    Humidity_Min = ''
    Humidity_Max = ''
    Humidity_Average = ''
    Humidity_Unit = ''

    AirDensity_Min = ''
    AirDensity_Max = ''
    AirDensity_Average = ''
    AirDensity_Unit = ''


class TMethods:
    def __init__(self, Name, Process, Description, MethodID):
        self.Name = Name
        self.Process = Process
        self.Description = Description
        self.MethodID = MethodID

    Name = ''
    Process = ''
    Version = ''
    Index = ''
    Description = ''
    MethodID = ''


# class TReferenceWeightSets:
#     def __init__(self, SerialNumber, Class, Range, Comment, CertificateNumber, CommonAlloyMaterialDensity):
#         self.SerialNumber = SerialNumber
#         self.Class = Class
#         self.Range = Range
#         self.Comment = Comment
#         self.CertificateNumber = CertificateNumber
#         self.CommonAlloyMaterialDensity = CommonAlloyMaterialDensity
#
#     SerialNumber = ''
#     Class = ''
#     Range = ''
#     CommonAlloyMaterial = ''
#     CommonAlloyMaterialDensity = ''
#     CommonAlloyMaterialDensityUnit = ''
#     CommonShape = ''
#     LastCalibrationDate = ''
#     CertificateNumber = ''
#     NextCalibrationDate = ''
#     Comment = ''
#
#
# class TReferenceWeights:
#     def __init__(self, Index, SerialNumber, NominalWeight, NominalWeightUnit, WeightId, Class, ConventionalMassError):
#         self.Index = Index
#         self.SerialNumber = SerialNumber
#         self.NominalWeight = NominalWeight
#         self.NominalWeightUnit = NominalWeightUnit
#         self.WeightId = WeightId
#         self.Class = Class
#         self.ConventionalMassError = ConventionalMassError
#
#     Index = ''
#     SerialNumber = ''
#     NominalWeight = ''
#     NominalWeightUnit = ''
#     WeightId = ''
#     Density = ''
#     DensityUnit = ''
#     DensityUncertainty = ''
#     DensityUncertaintyUnit = ''
#     Class = ''
#     ConventionalMass = ''
#     ConventionalMassUnit = ''
#     ConventionalMassError = ''
#     ConventionalMassErrorUnit = ''
#     TrueMass = ''
#     TrueMassUnit = ''
#     TrueMassError = ''
#     TrueMassErrorUnit = ''
#     AirDensityAtCalibration = ''
#     AirDensityAtCalibrationUnit = ''
#     AirTemperatureAtCalibration = ''
#     AirTemperatureUnitAtCalibration = ''
#     ExpandedMassErrorUncertainty = ''
#     ExpandedMassErrorUncertaintyUnit = ''
#     CertificateNumber = ''
#
#
# class TMassComparators:
#     def __init__(self, Index, Model, Name, SerialNumber, Description):
#         self.Index = Index
#         self.Model = Model
#         self.Name = Name
#         self.SerialNumber = SerialNumber
#         self.Description = Description
#
#     Index = ''
#     Model = ''
#     Name = ''
#     SerialNumber = ''
#     Description = ''
#     LastCalibrationDate = ''
#     NextCalibrationDate = ''
#     LastAdjustmentDate = ''
#     NextAdjustmentDate = ''
#
#
# class TMeasurementReadings:
#     SeriesIndex = ''
#     ProcessStepIndex = ''
#     WeightInstruction = ''
#     WeightReading = ''
#     WeightReadingUnit = ''
#     Step = ''
#
#
# class TTestWeightCalibrations:
#     AsReturned = True
#     Class = ''
#     PlusMinus = ''
#     Nominal = ''
#     NominalUnit = ''
#     WeightID = ''
#     Density = ''
#     DensityUnit = ''
#     DensityUncertainty = ''
#     DensityUncertaintyUnit = ''
#     Shape = ''
#     VolumeExpansionFactor = ''
#     VolumeExpansionFactorUnit = ''
#     VolumeExpansionFactorUncertainty = ''
#     VolumeExpansionFactorUncertaintyUnit = ''
#     AlloyOrMaterial = ''
#     CalibrationDate = ''
#     MassComparator = TMassComparators
#     ReferenceWeight = TReferenceWeights
#     CheckStandard = ''
#     SensitivityWeight = ''
#     AirDensityCalculation = ''
#     Method = TMethods
#     User = ''
#     TrueMassCorrection = ''
#     TrueMassCorrectionUnrounded = ''
#     TrueMassCorrectionUnit = ''
#     TrueMass = ''
#     TrueMassUnit = ''
#     ConventionalMassCorrection = ''
#     ConventionalMassCorrectionUnrounded = ''
#     ConventionalMassCorrectionUnit = ''
#     ConventionalMass = ''
#     ConventionalMassUnit = ''
#     MassVsBrassCorrection = ''
#     MassVsBrassCorrectionUnrounded = ''
#     MassVsBrassCorrectionUnit = ''
#     MassVsBrass = ''
#     MassVsBrassUnit = ''
#     CombinedMassUncertainty = ''
#     CombinedMassUncertaintyUnit = ''
#     ExpandedMassUncertainty = ''
#     ExpandedMassUncertaintyUnit = ''
#     ExpansionFactor = ''
#     Tolerance = ''
#     OneThirdTolerance = ''
#     ToleranceUnit = ''
#     CalibrationResult = ''
#     AirDensity = ''
#     AirDensityUnit = ''
#     AirTemperature = ''
#     AirTemperatureUncertainty = ''
#     AirTemperatureUnit = ''
#     AirTemperatureSensor = ''
#     AirPressure = ''
#     AirPressureUncertainty = ''
#     AirPressureUnit = ''
#     AirPressureSensor = ''
#     Humidity = ''
#     HumidityUnit = ''
#     HumidityUncertainty = ''
#     HumiditySensor = ''
#     MeasurementReadings = []
#

#
#
# class TTestWeightSet:
#     SerialNumber = ''
#     Manufacturer = ''
#     InternalID = ''
#     AccuracyClass = ''
#     MassDefinition = ''
#     NextCalibration = ''
#     Range = ''
#     CommonShape = ''
#     CommonAlloyMaterial = ''
#     CommonAlloyMaterialDensity = ''
#     CommonAlloyMaterialDensityUnit = ''
#     CalibratedBy = ''
#     Comment = ''
#     Description = ''
#     NextCalibrationDate = ''
#     RegNumber = ''
#     TestWeightCalibrations = []
#
#
# class TWeightSetCalibration:
#     StartDate = ''
#     EndDate = ''
#     CertificateNumber = ''
#     LevelConfidence = ''
#     CalibratedBy = ''
#     ContactOwner = TContactOwner
#     TestWeightSet = TTestWeightSet
#     EnvironmentalConditions = TEnvironmentalConditions
#     Methods = []
#     ReferenceWeightSets = []
#     ReferenceWeights = []
#     MassComparators = []
#
#
# class TWeightSetCalibrationExport:
#     Generated = ''
#     Language = ''
#     WeightSetCalibration = TWeightSetCalibration


# класс конвертора

class DemonConvertation(Thread):
    logfile = 'report.log'
    header = r'Клинский филиал ФБУ "ЦСМ Московской области", 141600, МО г. Клин, ул. Дзержининского, д.2'
    xml_loaded = bool
    runned = bool
    pathname = str(os.path.dirname(sys.argv[0])).replace('/', '\\')
    xml_folder = pathname
    Excel_folder = pathname
    template_filename = pathname + r"\шаблон.xls"
    TemplateApprovalReport = ""
    TemplateApprovalCert = ""
    TemplateError = ""
    TemplateProtocolCal = ""
    TemplateCalCert = ""
    TemplateComparatorApprovalProtocol = pathname + r'\templates\Протокол_поверки_компаратора.xls'

    autoopen = bool  # автоокрытие отчета после его создания
    autodelXML = bool  # автоудаление XML после создания отчета 
    ApprovalProtocolNum = 0  # номер протокол поверки
    ApprovalCertNum = 0  # номер свидетельства о поверке
    CalCertNum = 0  # номер калибровочного сертификата
    ErrorNum = 0  # номер извещения о непригодности 
    CalProtocolNum = 0  # номер протокола калибровки

    ApprovalProtocolEnable = bool  # признак создания протокола калибровки 
    ApprovalCertEnable = bool  # признак создания свидетельства о поверке
    ErrorEnable = bool  # признак создания извещения о непригодности
    CalProtocolEnable = bool  # признак создания протокола калибровки
    CalCertEnable = bool  # признак создания сертификата калибровки
    config_filename = 'config.ini'  # файл конфигурации
    conf = configparser.RawConfigParser()
    CSM = ''  # название ЦСМ 
    ApprovalProtocolFolder = 'Протоколы поверки'  # названия папок
    ApprovalCertFolder = 'Свидетельства о поверке'  # ...
    ErrorFolder = 'Извещения о непригодности'  # ...
    CalProtocolFolder = 'Протоколы калибровки'  # ...
    CalCertFolder = 'Сертификаты о калибровке'  # ...
    models = ""  # названия моделей компараторов

    # данные из XML
    xml = etree
    EndDate = ''  # дата поверки
    CI_Name = ''  # наименование СИ

    # Customer = TContactOwner()
    Company_Name = ''  # название заказчика
    CustomerNumber = ''  # номер заказчика
    CalibratedBy = ''  # ФИО поверителя
    HeadFIO = 'Писаревский Р.И.'
    HeadName = 'Начальник отдела мех. СИ'
    Adress = ''  # адрес заказчика
    Density = ''  # плотность материала гирь
    CertificateNumber = ''  # номер документа

    AirTemperature_Min = ''  # минимальная температура воздуха
    AirTemperature_Max = ''  # максимальная температура воздуха
    AirTemperature_Avr = ''  # средняя температура воздуха
    AirTemperature_Unit = ''  # размерность температуры
    Humidity_Min = ''  # минимальная влажность воздуха
    Humidity_Max = ''  # максимальная влажность воздуха
    Humidity_Avr = ''  # средняя влажность воздуха
    Humidity_Unit = ''  # размерность отностиельной влажности воздуха
    AirPressure_Min = ''  # минимальное давление воздуха
    AirPressure_Max = ''  # максимальное давление воздуха
    AirPressure_Avr = ''  # среднее давление воздуха
    AirPressure_Unit = ''  # размерность давления воздуха
    AirDensity_Min = ''  # минимальаня плотность воздуха
    AirDensity_Max = ''  # маскимальная плотность воздуха
    AirDensity_Avr = ''  # средняя плотность воздуха
    AirDensity_Unit = ''  # размерность плотности воздуха

    Method_Name = ''  # метод поверки
    Method_Process = ''  # название процесса поверки
    Method_ID = ''  # Признак метода Калибровка или Поверка и название метода через пробел

    MassComparators = []
    MassComparators_Info = ''  # информация о компараторах
    MassComparator_Model = list()  # модели компараторов
    MassComparator_SerialNumber = list()  # серийные номера компараторов
    MassComparator_Description = list()  # метрологические характеристики, описание компараторов

    ReferenceWeightSets = []
    ReferenceWeighs = []
    ReferenceWeightSet_SerialNumber = list()  # серийные номера эталонных наборов
    ReferenceWeightSet_Class = list()  # классы эталонных наборов
    ReferenceWeightSet_Range = list()  # диапазоны эталонных наборов
    ReferenceWeightSet_Comment = list()  # номер в реестре эталонов, информация об эталонных наборах
    ReferenceWeightSet_RegNumber = list()
    ReferenceWeightSet_NextCalibrationDate = list()
    ErrorResults = []

    TestWeights = []
    TestWeightSet_Manufacturer = ''  # производитель гирь
    TestWeightSet_AccuracyClass = ''  # класс точности испытуемого набора
    TestWeightSet_Range = ''  # диапазон номинальных заначений массы испытуемого набора
    TestWeightSet_SerialNumber = ''  # серийный номер испытуемого набора
    TestWeightSet_Description = ''  # год выпуска испытуемого набора
    TestWeight_CalibrationsCount = ''  # количество испытуемых гирь
    TestWeightSet_InternalID = ''  # номер клейма предыдущей поверки
    TestWeightSet_Comment = ''  # номер в госреестре
    Test_Passed = True

    TestWeight_Nominal = list()  # номиналы гирь
    TestWeight_NominalUnit = list()  # размерности гирь
    TestWeight_Tolerance = list()  # допуски гирь
    TestWeight_WeightId = list()

    ReferenceWeight_ConventionalMassError = list()  # погрешность эталона
    ReferenceWeight_ConventionalMassErrorUnit = list()
    ConventionalMassCorrection = list()  # отклонение от номинальной массы
    ConventionalMassCorrectionUnit = list()
    ConventionalMass = list()  # условная масса
    ConventionalMassUnit = list()  # размерность условной массы
    ExpandedMassUncertainty = list()  # расширенная неопределенность
    ExpandedMassUncertaintyUnit = list()  # размерность расширенной неопределенности
    MesurmentUnit = 'мг'
    TolUnit = 'мг'

    A1 = []
    A2 = []
    B1 = []
    B2 = []
    Diff = []
    Avr = []
    round_number = 5

    # конструктор класса конвертора
    def __init__(self):
        Thread.__init__(self)
        self.xml_loaded = False
        self.update_settings()

    # устанавливаем признаки протоколов
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

    # устанавливаем папку XML
    def setXmlFolder(self, xml_folder):
        self.conf.read(self.config_filename)
        self.conf.set('path', 'xml', xml_folder)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.xml_folder = xml_folder

    # устанавливаем папку Excel
    def setExcelFolder(self, Excel_folder):
        self.conf.read(self.config_filename)
        self.conf.set('path', 'Excel', Excel_folder)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.Excel_folder = Excel_folder
        self.folderExist()

    # проверяем существование папок
    def folderExist(self):
        if not os.access(self.Excel_folder, os.F_OK):
            os.mkdir(self.Excel_folder)
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

    # устанавливаем названия шаблонов
    def setTemplateFilename(self, template_filename, template):
        self.conf.read(self.config_filename)
        self.conf.set('path', template, template_filename)
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        if template == 'TemplateApprovalProtocol': self.TemplateApprovalReport = template_filename
        if template == 'TemplateApprovalCert': self.TemplateApprovalCert = template_filename
        if template == 'TemplateCalProtocol': self.TemplateCalReport = template_filename
        if template == 'TemplateCalCert':  self.TemplateCalCert = template_filename
        if template == 'TemplateError':  self.TemplateError = template_filename

    # устнавливаем признак автооокрытия
    def setAutoOpen(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autoopen', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.autoopen = set

    # устанавливаем признак автозапуска
    def setAutoStart(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autostart', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.runned = set

    # устанавливаем признак автоудаления XML
    def setAutoDelXML(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('auto', 'autodelXML', str(set))
        with open(self.config_filename, "w") as config:
            self.conf.write(config)
        self.autodelXML = set

    # устанавливаем имя ЦСМ
    def setNameCSM(self, set):
        self.conf.read(self.config_filename)
        self.conf.set('name', 'CSMName', str(set))
        with open(self.config_filename, 'w') as config:
            self.conf.write(config)
        self.CSM = set

    # устанавливаем номер документа
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

    # увличиваем номер документа
    def incNums(self, set, name):
        return set

    # Функция обновления настроек
    def update_settings(self):
        logging.debug('Обновление настроек')
        dir_path = os.path.dirname(sys.argv[0]).replace('/', '\\')
        self.conf.read(self.config_filename)
        xml_folder = self.conf.get('path', 'xml')
        if xml_folder[1] != ':': xml_folder = dir_path + "\\" + xml_folder
        Excel_folder = self.conf.get('path', 'Excel')
        if Excel_folder[1] != ':': Excel_folder = dir_path + "\\" + Excel_folder
        template_filename = self.conf.get('path', 'Template')
        autostart = self.conf.get('auto', 'autostart')
        autoopen = self.conf.get('auto', 'autoopen')
        autodelXML = self.conf.get('auto', 'autodelXML')
        CSM = self.conf.get('name', 'CSMName')
        self.HeadName = self.conf.get('FIO', 'headname')
        self.HeadFIO = self.conf.get('FIO', 'headfio')
        self.models = self.conf.get('comparators', 'models').split('\n')
        for i in range(0, len(self.models)):
            self.models[i] = str(self.models[i]).split(';')

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

        self.TemplateApprovalReport = self.conf.get('path', 'TemplateApprovalProtocol')
        if self.TemplateApprovalReport[
            1] != ':': self.TemplateApprovalReport = dir_path + "\\" + self.TemplateApprovalReport
        self.TemplateApprovalCert = self.conf.get('path', 'TemplateApprovalCert')
        if self.TemplateApprovalCert[1] != ':': self.TemplateApprovalCert = dir_path + "\\" + self.TemplateApprovalCert
        self.TemplateCalReport = self.conf.get('path', 'TemplateCalProtocol')
        if self.TemplateCalReport[1] != ':': self.TemplateCalReport = dir_path + "\\" + self.TemplateCalReport
        self.TemplateCalCert = self.conf.get('path', 'TemplateCalCert')
        if self.TemplateCalCert[1] != ':': self.TemplateCalCert = dir_path + "\\" + self.TemplateCalCert
        self.TemplateError = self.conf.get('path', 'TemplateError')
        if self.TemplateError[1] != ':': self.TemplateError = dir_path + "\\" + self.TemplateError
        # startpath = os.path.dirname(sys.argv[0]).replace('/','\\')
        if os.access(xml_folder, 0) != True:
            self.setXmlFolder(str(self.pathname) + r"\xml")

        if os.access(Excel_folder, 0) != True:
            self.setExcelFolder(str(self.pathname) + r"\Excel")

        if os.access(self.TemplateApprovalReport, 0) != True:
            self.setTemplateFilename(str(self.pathname) + r"\templates\Протокол_поверки_калибровки.docx",
                                     "TemplateApprovalProtocol")

        if os.access(self.TemplateApprovalCert, 0) != True:
            self.setTemplateFilename(str(self.pathname) + r"\templates\Свидетельство_о_поверке.docx",
                                     "TemplateApprovalCert")

        if os.access(self.TemplateError, 0) != True:
            self.setTemplateFilename(str(self.pathname) + r"\templates\Извещение_о_непригодности.docx", "TemplateError")

        if os.access(self.TemplateCalReport, 0) != True:
            self.setTemplateFilename(str(self.pathname) + r"\templates\Протокол_поверки_калибровки.docx",
                                     "TemplateCalProtocol")

        if os.access(self.TemplateCalCert, 0) != True:
            self.setTemplateFilename(str(self.pathname) + r"\templates\Сертификат_о_калибровке.docx", "TemplateCalCert")

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

        if self.CSM != "Клинский ЦСМ":
            self.header = self.CSM
        else:
            self.header = r'Клинский филиал ФБУ "ЦСМ Московской области", 141600, МО г. Клин, ул. Дзержининского, д.2'

    # функция запуска фонового процесса слежения за папкой xml_folder
    # если в этой папке появляется файл .xml то он автоматически конвертируется в протокол
    def run(self):
        self.update_settings()
        self.folderExist()
        logging.debug('Ожидание файла xml')
        while self.runned:
            file = os.listdir(self.xml_folder)
            if len(file) != 0:
                for i in file:
                    ext = i[-4:]
                    if ext == '.xml':
                        self.convertation(self.xml_folder + '\\' + i)
                        sleep(1)
            sleep(1)

    # перевод названия единиц измерения
    @staticmethod
    def correctUnit(unit):
        unit = str(unit).replace('ug', 'мкг')
        unit = str(unit).replace('mg', 'мг')
        unit = str(unit).replace('kg', 'кг')
        unit = str(unit).replace('g', 'г')
        return unit

    # округление строковых чисел
    @staticmethod
    def roundStr(_str, num):
        _str = str(_str).replace(',', '.')
        _str = float(_str)
        _str = round(_str, num)
        return str(_str).replace('.', ',')

    # проверка имени файла
    @staticmethod
    def rightFileName(_str):
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

    # формирование протокола поверки компаратора
    def ComparatorApprovalReport(self, xml_filename):
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
        # Method_ID = Method[0].text
        # CallReport = False
        # if str(Method_ID).find('Калибровка') != -1:
        #     CallReport = True

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
            Company_Name = Company  # название заказчика
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

        # Method_Name = Method[0].get('Name')  # метод поверки
        Method_Name = self.TestWeightSet_Description
        Reestr_Number = self.TestWeightSet_Comment
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

            rb = xlrd.open_workbook(self.TemplateComparatorApprovalProtocol, formatting_info=True,
                                    on_demand=True)  # открываем книгу
            wb = xlcopy(rb)  # копируем книгу в память

            # Печать протокола калибровки
            ws = wb.get_sheet(0)  # выбираем лист протокола поверки
            ws.footer_str = str('&LПротокол №' + str(self.ApprovalProtocolNum) + '&CСтраница &P из &N').encode('utf-8')
            ws.header_str = str(self.header).encode('utf-8')

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
                ws.write(1, 4, " ", styleCellBottomCenter)  # номер протокола

            ws.write(2, 1, self.EndDate)  # дата поверки

            ws.write(3, 1, CI_Name)  # наименование СИ

            ws.write(4, 1, self.MassComparator_SerialNumber[0])  # год выпуска
            ws.write(5, 1, self.Company_Name)  # название заказчика
            ws.write(6, 1, Reestr_Number)  # номер заказчика
            ws.write(7, 1, Method_Name)  # производитель гирь

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
                    _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    _ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                            _ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                row += 1

            for comp in MassComparator:
                # название компаратора
                MassComparator_Model = comp.get('Model')
                ws.write(3, 1, MassComparator_Model)
                MassComparator_SerialNumber = comp.get('SerialNumber')
                ws.write(4, 1, MassComparator_SerialNumber)
                MassComparator_Description = comp.find('Description').text
                # описание компаратора. В поле Описание (Description) должны быть записаны дискретность и СКО модели компаратора
                # ws.write(row, 4, MassComparator_Description, styleCellCenter)
                row += 1

            # TODO: Настройки в шаблон
            Row = 46
            RowMetr = 5
            RowWeightReading = Row

            for i in TestWeightCalibrationAsReturned:
                Nominal = i.find('Nominal').text  # type: str
                NominalUnit = i.find('NominalUnit').text
                if NominalUnit == 'g':
                    NominalUnit = 'г'
                if NominalUnit == 'mg':
                    NominalUnit = 'мг'
                if NominalUnit == 'kg':
                    NominalUnit = 'кг'

                Index = i.find('ReferenceWeight').text
                ReferenceWeight_ConventionalMassError = 0
                Tolerance = i.find('Tolerance').text
                for j in ReferenceWeight:
                    if Index == j.get('Index'):
                        ReferenceWeight_ConventionalMassError = j.get('ConventionalMassError')
                ws.write(Row, 0, str.strip(Nominal + NominalUnit), styleCellCenterSinglCellTop)
                MeasurementReadings = i.find('MeasurementReadings')
                try:
                    WeightID = i.find('WeightID').text
                    WeightReading = MeasurementReadings.findall('WeightReading')
                    round_number = 5
                    A1 = []
                    A2 = []
                    B1 = []
                    B2 = []
                    Diff = []
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
                                for y in range(2, 5):
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

                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))

                            diff = B1[cicle] - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            B2[cicle] = float(B2[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))

                            diff = (B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
                                     styleCellLeftBottom)
                            RowWeightReading += 1

                    Avr = round(mean(Diff), round_number)

                    ws.write(Row, 3, Avr, styleCellCenterSinglCellTop)

                    Std = round(stdev(Diff), round_number)

                    ws.write(Row, 4, Std, styleCellCenterSinglCellTop)

                    # if MassComparator_Model == 'XP505':
                    #     if float(Nominal) <= 10:  ws.write(Row, 5, 0.01, styleCellCenterSinglCellTop)
                    #     if (10 < float(Nominal)) and (float(Nominal) <= 200): ws.write(Row, 5, 0.02, styleCellCenterSinglCellTop)
                    #     if float(Nominal) > 200: ws.write(Row, 5, 0.035, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPE505C':
                    #     if float(Nominal) <= 20:  ws.write(Row, 5, 0.008, styleCellCenterSinglCellTop)
                    #     if (20 < float(Nominal)) and (float(Nominal) <= 200): ws.write(Row, 5, 0.015, styleCellCenterSinglCellTop)
                    #     if float(Nominal) > 200: ws.write(Row, 5, 0.03, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPE32003LC':
                    #     if float(Nominal) <= 2000:  ws.write(Row, 5, 5, styleCellCenterSinglCellTop)
                    #     if float(Nominal) > 2000: ws.write(Row, 5, 10, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'AX12004':
                    #     if float(Nominal) <= 2000:  ws.write(Row, 5, 0.2, styleCellCenterSinglCellTop)
                    #     if float(Nominal) > 2000: ws.write(Row, 5, 0.25, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XP2004S':
                    #     if float(Nominal) <= 2000:  ws.write(Row, 5, 0.1, styleCellCenterSinglCellTop)
                    #
                    #
                    # if MassComparator_Model == 'XP5003S':
                    #     if float(Nominal) <= 5000:  ws.write(Row, 5, 1, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPE26C':
                    #     if float(Nominal) <= 1:  ws.write(Row, 5, 0.0006, styleCellCenterSinglCellTop)
                    #     if (1 < float(Nominal)) and (float(Nominal) <= 20): ws.write(Row, 5, 0.0012, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XP26C':
                    #     if float(Nominal) <= 1:  ws.write(Row, 5, 0.001, styleCellCenterSinglCellTop)
                    #     if (1 < float(Nominal)) and (float(Nominal) <= 20): ws.write(Row, 5, 0.0015, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPR6U':
                    #     nom = float(Nominal.replace(',', '.'))
                    #     if nom <= 0.2:  ws.write(Row, 5, 0.00015, styleCellCenterSinglCellTop)
                    #     else:           ws.write(Row, 5, 0.00027, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPE56C':
                    #     nom = float(Nominal.replace(',', '.'))
                    #     if nom <= 1:  ws.write(Row, 5, 0.001, styleCellCenterSinglCellTop)
                    #     else:           ws.write(Row, 5, 0.003, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XPE2004SC':
                    #     nom = float(Nominal.replace(',', '.'))
                    #     if nom <= 2000:  ws.write(Row, 5, 0.1, styleCellCenterSinglCellTop)
                    #
                    # if MassComparator_Model == 'XP26003L':
                    #     nom = float(Nominal.replace(',', '.'))
                    #     if nom <= 1000:  ws.write(Row, 5, 2, styleCellCenterSinglCellTop)
                    #     else:           ws.write(Row, 5,  3, styleCellCenterSinglCellTop)

                    if WeightID != '':
                        ws.write(Row, 5, WeightID, styleCellCenterSinglCellTop)
                        Tol = float(str(WeightID).replace(',', '.'))

                        if Std <= Tol and Test_Passed != False:
                            Test_Passed = True
                        else:
                            Test_Passed = False

                    for x in range(Row + 1, Row + len(WeightReading)):
                        ws.write(x, 0, '', styleCellLeftLine)
                        ws.write(x, 5, '', styleCellRightLine)
                except:
                    logging.exception('Ошибка формирования протокола поверки компаратора')
                    logging.debug('Ошибка формирования протокола поверки компаратора')
                Row = RowWeightReading

            for y in range(0, 6):
                ws.write(Row, y, '', styleCellTopLine)

            if Test_Passed:
                ws.write(Row + 2, 0,
                         'Заключение по результатам поверки: компаратор признан пригодным к использованию')
                ws.write(Row + 4, 0, 'На основании результатов поверки выдан сертификат № ' + str(
                    self.CalCertNum) + '        от  ' + EndDate)
            else:
                ws.write(Row + 2, 0,
                         'Заключение: по результатам поверки компаратор признан непригодным к использованию')
                ws.write(Row + 4, 0,
                         'На основании результатов поверки выдано извещение о непригодности № _______________________ от _____._____________._______г.')

            ws.write(Row + 7, 0, 'Поверитель:_____________________ ' + CalibratedBy)
            ws.write(Row + 7, 6, 'Дата протокола: ' + str(datetime.today().strftime("%d.%m.%Y")))

            # ws.insert_bitmap('logo.bmp', 1, 7)

            # Составление сертификата калибровки

            NextYear = datetime.strptime(str(EndDate), "%d.%m.%Y").date()
            NextYear += relativedelta(years=1)

            if ReferenceWeightSet_RegNumber != "":
                ReferenceWeightSet_RegNumber = ", рег. № " + ReferenceWeightSet_RegNumber

            # Сохранение сертификата и протокола
            # сохранение данных в новый документ
            date_time = datetime.now().strftime("%d%m%Y_%H%M%S")

            file_to_save = self.rightFileName(
                "Протокол поверки компаратора " + MassComparator_Model + " заводской номер " + MassComparator_SerialNumber + "  " + date_time + '.xls')
            fileApprovalProtocol = self.Excel_folder + '\\' + self.ApprovalProtocolFolder + '\\' + file_to_save
            wb.save(fileApprovalProtocol)
            logging.debug(
                'Файл протокола поверки компаратора ' + MassComparator_Model + ' ' + MassComparator_SerialNumber + ' создан: ' + fileApprovalProtocol)
            if self.autoopen == True:
                os.startfile(fileApprovalProtocol)

    # формирование поверочного протокола Клин
    def ApprovalReportKlin(self, xml_filename):

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

            rb = xlrd.open_workbook(self.TemplateApprovalReport, formatting_info=True,
                                    on_demand=True)  # открываем книгу
            wb = xlcopy(rb)  # копируем книгу в память
            ws = wb.get_sheet(0)  # выбираем лист протокола поверки
            # ws.name = 'Клинский филиал ФБУ "ЦСМ Московской области", 141600, МО г. Клин, ул. Дзержининского, д.2'
            # 'Протокол №' + str(self.ApprovalProtocolNum).replace('/', '  ')
            ws.footer_str = str('&LПротокол №' + str(self.ApprovalProtocolNum) + '&CСтраница &P из &N').encode('utf-8')
            ws.header_str = str(self.header).encode(
                'utf-8')

            rbCert = xlrd.open_workbook(self.TemplateApprovalCert, formatting_info=True,
                                        on_demand=True)  # открываем книгу
            wbCert = xlcopy(rbCert)  # копируем книгу в память
            wsCert = wbCert.get_sheet(0)  # выбираем лист свидетельства
            wsCertMert = wbCert.get_sheet(0)  # выбираем лист метрологических характеристик

            # стиль ячеки выравнивание по центру
            styleCellCenter = xlwt.easyxf(
                'border: top thin, left thin, bottom thin, right thin; align: horiz center')
            # стиль ячейки выравнивание влево
            styleCellLeft = xlwt.easyxf('border: top thin, left thin, bottom thin; align: horiz left')
            styleCellLeftSinglCell = xlwt.easyxf(
                'border: top thin, left thin, bottom thin, right thin; align: horiz left')
            # styleCellCenterSinglCell = xlwt.easyxf('border: top thin, left thin, bottom thin, right thin; align: horiz center')
            styleCellCenterSinglCellTop = xlwt.easyxf(
                'border: top thin, left thin, right thin; align: horiz center')
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
                ws.write(1, 3, self.ApprovalProtocolNum)  # номер протокола
            ws.write(3, 5, self.EndDate)  # дата поверки
            ws.write(4, 1, self.Company_Name)  # название заказчика
            ws.write(5, 1, self.TestWeightSet_Range)  # номинальное заначение массы
            ws.write(6, 1, self.TestWeightSet_SerialNumber)  # серийный номер
            ws.write(7, 1, self.TestWeightSet_AccuracyClass)  # класс гири
            ws.write(8, 1, self.TestWeightSet_Comment)  # номер в госреестре
            ws.write(9, 1, self.Method_ID)  # метод

            ws.write(14, 2, Density)  # плотность материала гирь

            ws.write(14, 1, AirTemperature_Min, styleCellCenter)
            ws.write(15, 1, AirTemperature_Max, styleCellCenter)
            ws.write(16, 1, AirTemperature_Avr, styleCellCenter)

            ws.write(14, 3, Humidity_Min, styleCellCenter)
            ws.write(15, 3, Humidity_Max, styleCellCenter)
            ws.write(16, 3, Humidity_Avr, styleCellCenter)

            ws.write(14, 5, AirPressure_Min, styleCellCenter)
            ws.write(15, 5, AirPressure_Max, styleCellCenter)
            ws.write(16, 5, AirPressure_Avr, styleCellCenter)

            ws.write(14, 7, AirDensity_Min, styleCellCenter)
            ws.write(15, 7, AirDensity_Max, styleCellCenter)
            ws.write(16, 7, AirDensity_Avr, styleCellCenter)

            # TODO: Настройки в шаблон
            row = 20
            # TODO: Метрологические характеристики набора
            ReferenceWeightSet_info = ''
            for ref in ReferenceWeightSet:
                # название набора гирь
                if ref.get('Range') != "":
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    _ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                    ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                            _ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                            ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
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
            Row = 29
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
                ws.write(Row, 8, float(Tolerance.replace(',', '.')), styleCellCenterSinglCellTop)
                MeasurementReadings = i.find('MeasurementReadings')
                try:
                    WeightReading = MeasurementReadings.findall('WeightReading')
                    RowWeightReading = Row

                    A1 = []
                    A2 = []
                    B1 = []
                    B2 = []
                    Diff = []
                    round_number = 5
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))
                            diff = B1[cicle] - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            B2[cicle] = float(B2[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))

                            diff = (B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
                                     styleCellLeftBottom)
                            RowWeightReading += 1

                    Avr = round(mean(Diff), round_number)

                    ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
                    ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection, round_number)
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
                    ws.write(Row, 4, float(ReferenceWeight_ConventionalMassError.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 5, float(ConventionalMassCorrection.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 6, float(ConventionalMass.replace(',', '.')), styleCellCenterSinglCellTop)
                    ws.write(Row, 7, float(str(ExpandedMassUncertainty).replace(',', '.')),
                             styleCellCenterSinglCellTop)

                    wsCert.write(RowMetr, 11, WeightID + Nominal + NominalUnit, styleCellCenter)
                    wsCert.write(RowMetr, 12, "", styleCellCenter)
                    wsCert.write(RowMetr, 13, ConventionalMass + ConventionalMassUnit, styleCellCenter)
                    wsCert.write(RowMetr, 14, "", styleCellCenter)
                    wsCert.write(RowMetr, 15, ConventionalMassCorrection + ConventionalMassCorrectionUnit,
                                 styleCellCenter)
                    wsCert.write(RowMetr, 16, "", styleCellCenter)
                    wsCert.write(RowMetr, 17, ExpandedMassUncertainty, styleCellCenter)
                    wsCert.write(RowMetr, 18, "", styleCellLeftSinglCell)

                    for x in range(Row + 1, Row + len(WeightReading)):
                        ws.write(x, 0, '', styleCellLeftLine)
                        ws.write(x, 8, '', styleCellRightLine)
                    Row = RowWeightReading
                    RowMetr += 1
                except:
                    WeightReading = ""

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
                    ws.write(Row + 4, 0,
                             'На основании результатов поверки выдано извещение о непригодности № ' + str(
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
                self.setNums('0' + str(int(self.ApprovalCertNum[1:-3]) + 1) + '-' + str(datetime.today().year)[2:],
                             'ApprovalCertNum')
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
                wsCert.write(20, 2, ReferenceWeightSet_info + " " + ReferenceWeightSet_RegNumber)
                wsCert.write(28, 0, AirTemperature_Avr + " oC,  отностительная влажность " + Humidity_Avr + "%",
                             styleCellBottom)

                fileApprovalCert = self.Excel_folder + '\\' + self.ApprovalCertFolder + r'\Свидетельство о поверке ' + file_to_save
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
                try:
                    self.setNums('0' + str(int(self.ErrorNum[1:]) + 1),
                                 'ErrorNum')
                except:
                    logging.debug('Ошибка присвоения номер документа')
                # self.setNums(str(int(self.ErrorNum) + 1), 'ErrorNum')
                wsError.write(7, 2,
                              CI_Name + " " + TestWeightSet_Range + " " + TestWeightSet_AccuracyClass + " " + TestWeightSet_Comment,
                              styleCellBottom)
                wsError.write(13, 4, TestWeightSet_InternalID, styleCellBottom)
                wsError.write(15, 2, TestWeightSet_SerialNumber, styleCellBottom)
                wsError.write(17, 2, Company_Name + ", ИНН " + CustomerNumber, styleCellBottom)
                wsError.write(34, 0, str(datetime.today().strftime("%d.%m.%Y")), styleCellBottom)

                fileError = self.Excel_folder + '\\' + self.ErrorFolder + r'\Извещение о непригодности ' + file_to_save
                if self.ErrorEnable:
                    wbError.save(fileError)
                    if self.autoopen == True:
                        logging.debug('Извещение о непригодности для ' + self.Company + ' создано: ' + fileError)
                        os.startfile(fileError)
                if self.CalProtocolEnable:
                    wb.save(fileApprovalProtocol)
                    if self.autoopen == True:
                        logging.debug('Протокол поверки для ' + self.Company + ' создан: ' + fileApprovalProtocol)
                        os.startfile(fileApprovalProtocol)

    # формирование калибровочного протокола Клин
    def CalReportKlin(self, xml_filename):
        logging.debug('Формирование протокола калибровки для ' + self.CSM)
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

            rb = xlrd.open_workbook(self.TemplateCalReport, formatting_info=True, on_demand=True)  # открываем книгу
            wb = xlcopy(rb)  # копируем книгу в память
            ws = wb.get_sheet(0)  # выбираем лист протокола калибровки
            ws.footer_str = str('&LПротокол №' + str(self.ApprovalProtocolNum) + '&CСтраница &P из &N').encode('utf-8')
            ws.header_str = str(
                r'Клинский филиал ФБУ "ЦСМ Московской области", 141600, МО г. Клин, ул. Дзержининского, д.2').encode(
                'utf-8')
            ws.fit_width_to_pages = 1
            rbCert = xlrd.open_workbook(self.TemplateCalCert, formatting_info=True, on_demand=True)  # открываем книгу
            wbCert = xlcopy(rbCert)  # копируем книгу в память
            wsCert = wbCert.get_sheet(0)  # выбираем лист сертификата
            wsCertMert = wbCert.get_sheet(0)  # выбираем лист метрологических характеристик

            # стиль ячеки выравнивание по центру
            styleCellCenter = xlwt.easyxf(
                'border: top thin, left thin, bottom thin, right thin; align: horiz center')
            # стиль ячейки выравнивание влево
            styleCellLeft = xlwt.easyxf('border: top thin, left thin, bottom thin; align: horiz left')
            styleCellLeftSinglCell = xlwt.easyxf(
                'border: top thin, left thin, bottom thin, right thin; align: horiz left')
            # styleCellCenterSinglCell = xlwt.easyxf('border: top thin, left thin, bottom thin, right thin; align: horiz center')
            styleCellCenterSinglCellTop = xlwt.easyxf(
                'border: top thin, left thin, right thin; align: horiz center')
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
                ws.write(1, 3, self.CalProtocolNum)  # номер протокола
            ws.write(3, 5, self.EndDate)  # дата поверки
            ws.write(4, 1, self.Company_Name)  # название заказчика
            ws.write(5, 1, self.TestWeightSet_Range)  # номинальное заначение массы
            ws.write(6, 1, self.TestWeightSet_SerialNumber)  # серийный номер
            ws.write(7, 1, self.TestWeightSet_AccuracyClass)  # класс гири
            ws.write(8, 1, self.TestWeightSet_Comment)  # номер в госреестре
            ws.write(9, 1, str(self.Method_ID).strip('Калибровка '))  # метод

            # ws.write(14, 2, Density)  # плотность материала гирь

            ws.write(14, 1, AirTemperature_Min, styleCellCenter)
            ws.write(15, 1, AirTemperature_Max, styleCellCenter)
            ws.write(16, 1, AirTemperature_Avr, styleCellCenter)

            ws.write(14, 3, Humidity_Min, styleCellCenter)
            ws.write(15, 3, Humidity_Max, styleCellCenter)
            ws.write(16, 3, Humidity_Avr, styleCellCenter)

            ws.write(14, 5, AirPressure_Min, styleCellCenter)
            ws.write(15, 5, AirPressure_Max, styleCellCenter)
            ws.write(16, 5, AirPressure_Avr, styleCellCenter)

            ws.write(14, 7, AirDensity_Min, styleCellCenter)
            ws.write(15, 7, AirDensity_Max, styleCellCenter)
            ws.write(16, 7, AirDensity_Avr, styleCellCenter)

            # TODO: Настройки в шаблон
            row = 20
            # TODO: Метрологические характеристики набора
            ReferenceWeightSet_info = ''
            for ref in ReferenceWeightSet:
                # название набора гирь
                if ref.get('Range') != "":
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    _ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                    ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                            _ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                            ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
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
            Row = 29
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
                ws.write(Row, 8, float(Tolerance.replace(',', '.')), styleCellCenterSinglCellTop)
                MeasurementReadings = i.find('MeasurementReadings')
                try:
                    WeightReading = MeasurementReadings.findall('WeightReading')
                    RowWeightReading = Row

                    A1 = []
                    A2 = []
                    B1 = []
                    B2 = []
                    Diff = []
                    round_number = 5
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))
                            diff = B1[cicle] - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            B2[cicle] = float(B2[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))

                            diff = (B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
                                     styleCellLeftBottom)
                            RowWeightReading += 1

                    Avr = round(mean(Diff), round_number)

                    ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
                    ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection, round_number)
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
                    ws.write(Row, 4, float(ReferenceWeight_ConventionalMassError.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 5, float(ConventionalMassCorrection.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 6, float(ConventionalMass.replace(',', '.')), styleCellCenterSinglCellTop)
                    ws.write(Row, 7, float(str(ExpandedMassUncertainty).replace(',', '.')),
                             styleCellCenterSinglCellTop)

                    wsCert.write(RowMetr, 11, WeightID + Nominal + NominalUnit, styleCellCenter)
                    wsCert.write(RowMetr, 12, "", styleCellCenter)
                    wsCert.write(RowMetr, 13, ConventionalMass + ConventionalMassUnit, styleCellCenter)
                    wsCert.write(RowMetr, 14, "", styleCellCenter)
                    wsCert.write(RowMetr, 15, ConventionalMassCorrection + ConventionalMassCorrectionUnit,
                                 styleCellCenter)
                    wsCert.write(RowMetr, 16, "", styleCellCenter)
                    wsCert.write(RowMetr, 17, ExpandedMassUncertainty, styleCellCenter)
                    wsCert.write(RowMetr, 18, "", styleCellLeftSinglCell)

                    for x in range(Row + 1, Row + len(WeightReading)):
                        ws.write(x, 0, '', styleCellLeftLine)
                        ws.write(x, 8, '', styleCellRightLine)
                    Row = RowWeightReading
                    RowMetr += 1
                except:
                    WeightReading = ""

            for y in range(0, 9):
                ws.write(Row, y, '', styleCellTopLine)

            ws.write(Row + 2, 0,
                     'Заключение по результатам калибровки: гири пригодны к использованию по  классу точности ' + TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
            ws.write(Row + 4, 0, 'На основании результатов калибровки выдан сертификат о калибровке № ' + str(
                self.CalCertNum) + ' от _____._____________._______г.')

            ws.write(Row + 7, 0, 'Поверитель:_____________________ ' + CalibratedBy)
            ws.write(Row + 7, 6, 'Дата протокола: ' + str(datetime.today().strftime("%d.%m.%Y")))

            ws.insert_bitmap('logo.bmp', 1, 7)

            # сохранение данных в новый документ
            date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
            file_to_save = self.rightFileName(
                Company_Name.strip() + ' ' + TestWeightSet_AccuracyClass.strip() + ' ' + TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.xls')

            # Составление сертификата о калибровке

            fileCalProtocol = self.Excel_folder + '\\' + self.CalProtocolFolder + '\\' + file_to_save
            fileCalProtocol = fileCalProtocol.replace(',', ' ')

            NextYear = datetime.strptime(str(EndDate), "%d.%m.%Y").date()
            NextYear += relativedelta(years=1)
            wsCert.write(2, 4, self.CalCertNum, styleCellBottom)
            # self.setNums('0' + str(int(self.CalCertNum[1:-3]) + 1) + '-' + str(datetime.today().year)[2:], 'CalCertNum')
            wsCert.write(4, 8, str(NextYear.strftime("%d.%m.%Y")), styleCellBottom)
            wsCert.write(6, 2,
                         CI_Name + " " + TestWeightSet_Range + " " + TestWeightSet_AccuracyClass + " " + TestWeightSet_Comment,
                         styleCellBottom)
            wsCert.write(12, 4, TestWeightSet_InternalID,
                         styleCellBottom)  # серия и номер клейма предыдущей поверки
            # wsCert.write(10, 2, TestWeightSet_SerialNumber, styleCellBottom)
            wsCert.write(14, 2, TestWeightSet_SerialNumber, styleCellBottom)
            # wsCert.write(14, 1, Company_Name + ", ИНН " + CustomerNumber, styleCellBottom)
            Method_ID = str(Method_ID).strip('Калибровка')
            wsCert.write(16, 2, Method_ID, styleCellBottom)
            if ReferenceWeightSet_RegNumber != "":
                ReferenceWeightSet_RegNumber = ", рег.№ " + ReferenceWeightSet_RegNumber
            wsCert.write(20, 2, ReferenceWeightSet_info + " " + ReferenceWeightSet_RegNumber)
            wsCert.write(28, 0, AirTemperature_Avr + " oC,  отностительная влажность " + Humidity_Avr + "%",
                         styleCellBottom)

            fileCalCert = self.Excel_folder + '\\' + self.CalCertFolder + r'\Сертификат ' + file_to_save
            fileCalCert = fileCalCert.replace(',', ' ')

            if self.CalCertEnable:
                wb.save(fileCalCert)
                if self.autoopen == True:
                    logging.debug('Сертификат калибровки для ' + self.CSM + ' создан: ' + fileCalCert)
                    os.startfile(fileCalCert, 'open')

            if self.CalProtocolEnable:
                wb.save(fileCalProtocol)
                if self.autoopen == True:
                    logging.debug('Протокол калибровки для ' + self.CSM + ' создан: ' + fileCalProtocol)
                    os.startfile(fileCalProtocol)

    # формирование поверочного протокола
    def ApprovalReport(self, xml_filename=None):
        # Заполняем протокол поверки если есть положительные результаты AsReturned
        logging.debug('Формирование поверочного протокола для ' + self.Company)
        try:
            if len(self.TestWeight_Nominal) > 0:
                # открываем книгу
                rb = xlrd.open_workbook(self.TemplateApprovalReport,
                                        formatting_info=True,
                                        on_demand=True)
                wb = xlcopy(rb)  # копируем книгу в память
                ws = wb.get_sheet(0)  # выбираем лист протокола поверки
                # открываем книгу
                rbCert = xlrd.open_workbook(self.TemplateApprovalCert,
                                            formatting_info=True,
                                            on_demand=True)
                wbCert = xlcopy(rbCert)  # копируем книгу в память
                wsCert = wbCert.get_sheet(0)  # выбираем лист свидетельства
                wsCertMert = wbCert.get_sheet(0)  # выбираем лист метрологических характеристик

                # стиль ячеки выравнивание по центру
                styleCellCenter = xlwt.easyxf(
                    'border: top thin, left thin, bottom thin, right thin; align: horiz center')
                # стиль ячейки выравнивание влево
                styleCellLeft = xlwt.easyxf('border: top thin, left thin, bottom thin; align: horiz left')
                styleCellLeftSinglCell = xlwt.easyxf(
                    'border: top thin, left thin, bottom thin, right thin; align: horiz left')
                styleCellCenterSinglCellTop = xlwt.easyxf(
                    'border: top thin, left thin, right thin; align: horiz center')
                styleCellTopLine = xlwt.easyxf('border: top thin')
                styleCellLeftLine = xlwt.easyxf('border: left thin')
                styleCellRightLine = xlwt.easyxf('border: right thin')
                styleCellBottom = xlwt.easyxf('border: bottom thin')
                styleCellBorder = xlwt.easyxf('border: left thin, right thin')
                styleCellLeftBottom = xlwt.easyxf('border: left thin, bottom thin, right thin; align: horiz left')

                # КрасЦСМ номер протокола не печатаем
                if self.CSM != "КрасЦСМ":
                    ws.write(1, 4, self.ApprovalProtocolNum)  # номер протокола
                    self.setNums(self.ApprovalProtocolNum, 'ApprovalProtocolNum')
                ws.write(2, 1, self.EndDate)  # дата поверки
                ws.write(3, 1, self.CI_Name)  # наименование СИ
                ws.write(2, 6, self.TestWeightSet_AccuracyClass)  # класс точности
                ws.write(3, 6, self.TestWeightSet_Range)  # номинальное заначение массы
                ws.write(4, 6, self.TestWeightSet_SerialNumber)  # серийный номер
                ws.write(4, 1, self.TestWeightSet_Description)  # год выпуска
                ws.write(5, 1, self.Company_Name)  # название заказчика
                ws.write(6, 1, self.CustomerNumber)  # номер заказчика
                ws.write(7, 1, self.TestWeightSet_Manufacturer)  # производитель гирь
                ws.write(9, 1, self.TestWeightSet_InternalID)  # серия и номер клейма предыдущей поверки

                ws.write(14, 2, self.Density)  # плотность материала гирь

                ws.write(31, 1, self.AirTemperature_Min, styleCellCenter)
                ws.write(32, 1, self.AirTemperature_Max, styleCellCenter)
                ws.write(33, 1, self.AirTemperature_Avr, styleCellCenter)

                ws.write(31, 3, self.Humidity_Min, styleCellCenter)
                ws.write(32, 3, self.Humidity_Max, styleCellCenter)
                ws.write(33, 3, self.Humidity_Avr, styleCellCenter)

                ws.write(31, 5, self.AirPressure_Min, styleCellCenter)
                ws.write(32, 5, self.AirPressure_Max, styleCellCenter)
                ws.write(33, 5, self.AirPressure_Avr, styleCellCenter)

                ws.write(31, 7, self.AirDensity_Min, styleCellCenter)
                ws.write(32, 7, self.AirDensity_Max, styleCellCenter)
                ws.write(33, 7, self.AirDensity_Avr, styleCellCenter)

                row = 37
                n = len(self.ReferenceWeightSet_SerialNumber)
                i = 0
                while i < n:
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    ws.write(row, 2, self.ReferenceWeightSet_SerialNumber[i], styleCellCenter)
                    ws.write(row, 4, self.ReferenceWeightSet_Class[i] + ': ' + self.ReferenceWeightSet_Range[i],
                             styleCellCenter)
                    i += 1
                    row += 1
                n = len(self.MassComparator_SerialNumber)
                i = 0
                while i < n:
                    ws.write(row, 0, 'Компаратор массы ' + self.MassComparator_Model[i], styleCellCenter)
                    ws.write(row, 2, self.MassComparator_SerialNumber[i], styleCellCenter)
                    ws.write(row, 4, self.MassComparator_Description[i], styleCellCenter)
                    i += 1
                    row += 1

                # TODO: Настройки в шаблон

                Row = 46
                RowMetr = 5
                ReferenceWeightSet_info = ''
                # TODO: Настройки в шаблон
                n = len(self.TestWeight_Nominal)
                i = 0
                while i < n:
                    Nominal = self.TestWeight_Nominal[i]
                    NominalUnit = self.TestWeight_NominalUnit[i]
                    WeightID = self.TestWeight_WeightId[i]

                    if self.TestWeight_CalibrationsCount == '1':
                        ws.write(Row, 0, str.strip(Nominal + NominalUnit),
                                 styleCellCenterSinglCellTop)
                    else:
                        ws.write(Row, 0, str.strip(WeightID + Nominal + NominalUnit),
                                 styleCellCenterSinglCellTop)
                    ws.write(Row, 8, self.TestWeight_Tolerance[i], styleCellCenterSinglCellTop)
                    try:
                        RowWeightReading = Row
                        A1 = self.A1[i]
                        A2 = self.A2[i]
                        B1 = self.B1[i]
                        B2 = self.B2[i]
                        Diff = self.Diff[i]
                        # Определение метода
                        for cicle in range(int(len(A1))):
                            for x in range(RowWeightReading, RowWeightReading + 4):
                                for y in range(2, 8):
                                    ws.write(x, y, '', styleCellBorder)
                            ws.write(RowWeightReading, 1, str(cicle + 1) + ' A1 ' + str(A1[cicle]),
                                     styleCellLeftSinglCell)
                            RowWeightReading += 1
                            ws.write(RowWeightReading, 1, str(cicle + 1) + ' B1 ' + str(B1[cicle]),
                                     styleCellLeftSinglCell)
                            RowWeightReading += 1
                            if (len(B2) > 0):
                                ws.write(RowWeightReading, 1, str(cicle + 1) + ' B2 ' + str(B2[cicle]),
                                         styleCellLeftSinglCell)
                                RowWeightReading += 1
                            ws.write(RowWeightReading, 1, str(cicle + 1) + ' A2 ' + str(A2[cicle]),
                                     styleCellLeftSinglCell)
                            round_number = abs(
                                decimal.Decimal(str(A1[cicle]).replace(',', '.')).as_tuple().exponent) + 1
                            ws.write(RowWeightReading, 2, Diff[cicle], styleCellLeftBottom)
                            RowWeightReading += 1

                        ConventionalMassCorrection = self.ConventionalMassCorrection[i]
                        ConventionalMassCorrectionUnit = self.ConventionalMassCorrectionUnit[i]
                        ConventionalMass = self.ConventionalMass[i]
                        ConventionalMassUnit = self.ConventionalMassUnit[i]

                        ExpandedMassUncertainty = self.ExpandedMassUncertainty[i]
                        ExpandedMassUncertaintyUnit = self.ExpandedMassUncertaintyUnit[i]
                        Avr = self.Avr[i]
                        ReferenceWeight_ConventionalMassError = self.ReferenceWeight_ConventionalMassError[i]
                        ws.write(Row, 3, Avr, styleCellCenterSinglCellTop)
                        ws.write(Row, 4, ReferenceWeight_ConventionalMassError, styleCellCenterSinglCellTop)
                        ws.write(Row, 5, ConventionalMassCorrection, styleCellCenterSinglCellTop)
                        ws.write(Row, 6, ConventionalMass, styleCellCenterSinglCellTop)
                        ws.write(Row, 7, ExpandedMassUncertainty, styleCellCenterSinglCellTop)

                        for x in range(Row + 1, RowWeightReading):
                            ws.write(x, 0, '', styleCellLeftLine)
                            ws.write(x, 8, '', styleCellRightLine)
                        Row = RowWeightReading
                        RowMetr += 1
                        i += 1
                    except:
                        logging.debug('Ошибка формирования поверочного протокола')
                        WeightReading = ""

                for y in range(0, 9):
                    ws.write(Row, y, '', styleCellTopLine)

                if self.Test_Passed:
                    ws.write(Row + 2, 0,
                             'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + self.TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                    ws.write(Row + 4, 0, 'На основании результатов поверки выдано свидетельство о поверке № ' + str(
                        self.ApprovalCertNum) + ' от _____._____________._______г.')
                else:
                    if self.CSM != 'КрасЦСМ':
                        ws.write(Row + 2, 0,
                                 'Заключение по результатам поверки: гири не пригодны к использованию по классу точности ' + self.TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                        ws.write(Row + 4, 0,
                                 'На основании результатов поверки выдано извещение о непригодности № ' + str(
                                     self.ErrorNum) + ' от _____._____________._______г.')
                    else:
                        ws.write(Row + 2, 0,
                                 'Заключение по результатам поверки: гири пригодны к использованию по  классу точности ' + self.TestWeightSet_AccuracyClass + ' согласно ГОСТ OIML R111-1-2009')
                        ws.write(Row + 4, 0,
                                 'На основании результатов поверки выдано свидетельство о поверке № _______________________ от _____._____________._______г.')

                    pass

                ws.write(Row + 7, 0, 'Поверитель:_____________________ ' + self.CalibratedBy)
                ws.write(Row + 7, 6, 'Дата протокола: ' + str(datetime.today().strftime("%d.%m.%Y")))

                ws.insert_bitmap('logo.bmp', 1, 7)

                # сохранение данных в новый документ
                date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
                file_to_save = self.rightFileName(self.TestWeightSet_Comment + ' '
                                                  + self.Company_Name + ' '
                                                  + self.TestWeightSet_AccuracyClass + ' '
                                                  + self.TestWeightSet_SerialNumber + ' '
                                                  + date_time + '.xls')

                # Составление свидетельства о поверке

                fileApprovalProtocol = self.Excel_folder + '\\' + self.ApprovalProtocolFolder + '\\' + file_to_save
                fileApprovalProtocol = fileApprovalProtocol.replace(',', ' ')

                NextYear = datetime.strptime(str(self.EndDate), "%d.%m.%Y").date()
                NextYear += relativedelta(years=1)
                wsCert.write(2, 4, self.ApprovalCertNum, styleCellBottom)
                self.setNums('0' + str(int(self.ApprovalCertNum[1:-3]) + 1) + '-' + str(datetime.today().year)[2:],
                             'ApprovalCertNum')
                Method_ID = str(self.Method_ID).strip('Поверка')
                if self.ReferenceWeightSet_RegNumber != "":
                    ReferenceWeightSet_RegNumber = ', рег.№ ' + self.ReferenceWeightSet_RegNumber
                if self.ApprovalProtocolEnable == True:
                    wb.save(fileApprovalProtocol)
                    if self.autoopen == True:
                        logging.debug(
                            'Файл поверочного протокола для ' + self.Company + ' создан: ' + fileApprovalProtocol)
                        os.startfile(fileApprovalProtocol, 'open')
        except:
            logging.exception('Ошибка формирования поверочного протокола')

    # формирование протокола калибровки
    def CalReport(self, xml_filename):
        logging.debug('Формирование протокола калибровки для ' + self.Company)
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

            rb = xlrd.open_workbook(self.TemplateCalReport, formatting_info=True, on_demand=True)  # открываем книгу
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

                try:
                    self.setNums(str(int(self.CalProtocolNum[0:-3]) + 1) + '-' + str(datetime.today().year)[2:],
                                 'CalProtocolNum')
                except:
                    logging.debug('Ошибка установки номера документа')

            ws.write(2, 1, EndDate)  # дата поверки

            ws.write(3, 1, CI_Name)  # наименование СИ

            ws.write(2, 6, TestWeightSet_AccuracyClass)  # класс точности
            ws.write(3, 6, TestWeightSet_Range)  # номинальное заначение массы
            ws.write(4, 6, TestWeightSet_SerialNumber)  # серийный номер

            ws.write(4, 1, TestWeightSet_Description)  # год выпуска
            ws.write(5, 1, Company_Name)  # название заказчика
            ws.write(6, 1, CustomerNumber)  # номер заказчика
            ws.write(7, 1, TestWeightSet_Manufacturer)  # производитель гирь

            ws.write(14, 2, Density)  # плотность

            ws.write(31, 1, AirTemperature_Min, styleCellCenter)  # минимальная температура
            ws.write(32, 1, AirTemperature_Max, styleCellCenter)  # максимальная температура
            ws.write(33, 1, AirTemperature_Avr, styleCellCenter)  # средняя температура

            ws.write(31, 3, Humidity_Min, styleCellCenter)  # минимальная влажность
            ws.write(32, 3, Humidity_Max, styleCellCenter)  # максимальная влажность
            ws.write(33, 3, Humidity_Avr, styleCellCenter)  # средняя влажность

            ws.write(31, 5, AirPressure_Min, styleCellCenter)  # минимальное давление
            ws.write(32, 5, AirPressure_Max, styleCellCenter)  # максимальное давление
            ws.write(33, 5, AirPressure_Avr, styleCellCenter)  # среднее давление

            ws.write(31, 7, AirDensity_Min, styleCellCenter)  # минимальная плотность воздуха
            ws.write(32, 7, AirDensity_Max, styleCellCenter)  # максимальная плотность воздуха
            ws.write(33, 7, AirDensity_Avr, styleCellCenter)  # средняя плотность воздуха

            ReferenceWeightSet_info = ""
            # TODO: Настройки в шаблон
            row = 37
            # TODO: Метрологические характеристики набора
            for ref in ReferenceWeightSet:
                # название набора гирь
                if ref.get('Range') != "":
                    ws.write(row, 0, 'Набор гирь', styleCellCenter)
                    _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                    _ReferenceWeightSet_info = ref.get('Class') + ": " + ref.get('Range')
                    ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                    ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                    ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
                else:
                    for singleWeight in ReferenceWeight:
                        if singleWeight.get('SerialNumber') == ref.get('SerialNumber'):
                            ws.write(row, 0, 'Гиря', styleCellCenter)
                            _ReferenceWeightSet_SerialNumber = ref.get('SerialNumber')
                            _ReferenceWeightSet_info = ref.get('Class') + ": " + singleWeight.get(
                                'NominalWeight') + singleWeight.get('NominalWeightUnit')
                            ws.write(row, 4, _ReferenceWeightSet_info, styleCellCenter)
                            ws.write(row, 2, _ReferenceWeightSet_SerialNumber, styleCellCenter)
                            ReferenceWeightSet_info += _ReferenceWeightSet_info + " заводской номер " + _ReferenceWeightSet_SerialNumber
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
                ws.write(Row, 8, float(str(Tolerance).replace(',', '.')), styleCellCenterSinglCellTop)
                MeasurementReadings = i.find('MeasurementReadings')
                try:
                    WeightReading = MeasurementReadings.findall('WeightReading')
                    RowWeightReading = Row

                    A1 = []
                    A2 = []
                    B1 = []
                    B2 = []
                    Diff = []
                    round_number = 5
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))
                            diff = B1[cicle] - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
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
                            round_number = abs(decimal.Decimal(A1[cicle].replace(',', '.')).as_tuple().exponent) + 1
                            A1[cicle] = float(A1[cicle].replace(',', '.'))
                            B1[cicle] = float(B1[cicle].replace(',', '.'))
                            B2[cicle] = float(B2[cicle].replace(',', '.'))
                            A2[cicle] = float(A2[cicle].replace(',', '.'))
                            diff = (B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2
                            Diff.append(diff)
                            ws.write(RowWeightReading, 2, round(diff, round_number),
                                     styleCellLeftBottom)
                            RowWeightReading += 1

                    Avr = round(mean(Diff), round_number)

                    ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
                    ConventionalMassCorrection = self.roundStr(ConventionalMassCorrection, round_number)
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
                    ReferenceWeight_ConventionalMassError = self.roundStr(ReferenceWeight_ConventionalMassError,
                                                                          round_number)
                    ws.write(Row, 4, float(ReferenceWeight_ConventionalMassError.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 5, float(ConventionalMassCorrection.replace(',', '.')),
                             styleCellCenterSinglCellTop)
                    ws.write(Row, 6, float(ConventionalMass.replace(',', '.')), styleCellCenterSinglCellTop)
                    ws.write(Row, 7, float(str(ExpandedMassUncertainty).replace(',', '.')), styleCellCenterSinglCellTop)

                    wsCert.write(RowMetr, 11, WeightID + Nominal + NominalUnit, styleCellCenter)
                    wsCert.write(RowMetr, 12, "", styleCellCenter)
                    wsCert.write(RowMetr, 13, ConventionalMass + ConventionalMassUnit, styleCellCenter)
                    wsCert.write(RowMetr, 14, "", styleCellCenter)
                    wsCert.write(RowMetr, 15, ConventionalMassCorrection + ConventionalMassCorrectionUnit,
                                 styleCellCenter)
                    wsCert.write(RowMetr, 16, "", styleCellCenter)
                    wsCert.write(RowMetr, 17, ExpandedMassUncertainty, styleCellCenter)
                    wsCert.write(RowMetr, 18, "", styleCellLeftSinglCell)
                    RowMetr += 1
                    for x in range(Row + 1, Row + len(WeightReading)):
                        ws.write(x, 0, '', styleCellLeftLine)
                        ws.write(x, 8, '', styleCellRightLine)
                    # Row = RowWeightReading
                    Row = RowWeightReading
                    RowMetr += 1

                except:
                    WeightReading = ""

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
            try:
                self.setNums('0' + str(int(self.CalCertNum[1:]) + 1),
                             'CalCertNum')
            except:
                logging.debug('Ошибка установки номера документа')

            # self.setNums(str(int(self.CalCertNum) + 1), 'CalCertNum')

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
            fileCalCert = self.Excel_folder + '\\' + self.CalCertFolder + r'\Сертификат о калибровке ' + file_to_save

            if self.CalCertEnable == False:
                wb.save(fileCalProtocol)
                if self.autoopen == True:
                    logging.debug('Файл протокола клибровки для ' + self.Company + ' создан: ' + fileCalProtocol)
                    os.startfile(fileCalProtocol)
            else:
                wb.save(fileCalProtocol)
                wbCert.save(fileCalCert)
                if self.autoopen == True:
                    logging.debug('Файл протокола клибровки для ' + self.Company + ' создан: ' + fileCalProtocol)
                    logging.debug('Файл сертификата о клибровке для ' + self.Company + ' создан: ' + fileCalCert)
                    os.startfile(fileCalProtocol)
                    os.startfile(fileCalCert)

    # формирование свидетельства о поверке в формате docx
    def ApprovalCertDoc(self):
        logging.debug('Формирование свидетельства о поверке для ' + self.Company)
        # self.ApprovalCertNum = ''  # Для Тульского ЦСМ '     /10-02'
        DocNumber = str(self.ApprovalCertNum)
        # указываем путь до шаблона
        if self.TestWeightSet_Comment == '':
            template = self.TemplateApprovalCert
        else:
            template = self.pathname + r"\templates\Свидетельство_о_поверке_А5_эталон.docx"

        if self.TestWeightSet_InternalID == "":
            LastKleymo = '-'
        else:
            LastKleymo = self.TestWeightSet_InternalID
        EndDate = datetime.strptime(str(self.EndDate), "%d.%m.%Y").date()
        NextYear = EndDate
        NextYear = NextYear + relativedelta(years=+1, days=-1)
        # создаем объект и смотрим на имеющиемся поля
        document = MailMerge(template)

        months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', "июля", "августа", "сентября", "октября",
                  "ноября", "декабря"]

        ReferenseNumbers = ''

        MinRef = '1g'
        MaxRef = '50g'

        min = 5000
        max = 0
        for r in self.ReferenceWeighs:
            u = 1
            unit = r['NominalWeightUnit']
            if unit == 'kg': u = 1000
            if unit == 'g': u = 1
            if unit == 'mg': u = 0.001
            nominal = float(str(r['NominalWeight']).replace(',', '.')) * u
            if nominal < min:
                min = nominal
            if nominal >= max:
                max = nominal

        MaxRef = str(max) + ' г'
        MinRef = str(min) + ' г'

        ApprovedWith = 'в полном объеме'

        ReferenseNumbers = ''
        i = 1
        Etalon1 = ''
        Etalon2 = ''
        for r in self.ReferenceWeightSets:
            if i == 1: Etalon1 = r['Comment']
            if i == 2: Etalon1 += ', ' + r['Comment']
            if i == 3: Etalon2 = r['Comment']
            if i == 4: Etalon2 += ', ' + r['Comment']
            i += 1

        MaxLen1 = 45
        MaxLen2 = 63
        Name1 = ''
        Name2 = ''

        Name = self.TestWeightSet_Description.replace('\n', ' ')
        if len(Name) > MaxLen1:
            tmpstr = Name.split(' ')
            for i in tmpstr:
                if len(Name1 + i) < MaxLen1 and Name2 == '':
                    Name1 += i + ' '
                elif len(Name2 + i) < MaxLen2:
                    Name2 += i + ' '
        else:
            Name1 = Name

        Razrad = ''
        if (self.TestWeightSet_AccuracyClass == 'E2'):
            Razrad = '1'
        if (self.TestWeightSet_AccuracyClass == 'F1'):
            Razrad = '2'
        if (self.TestWeightSet_AccuracyClass == 'F2'):
            Razrad = '3'

        document.merge(
            RegNum=str(self.TestWeightSet_Comment),
            Razrad=str(Razrad),
            DocNumber=DocNumber,
            DayUntil=str(datetime.strftime(NextYear, "%d")),
            MounthUntil=str(months[NextYear.month - 1]),
            YearUntil=str(datetime.strftime(NextYear, "%Y")),
            DayCal=str(datetime.strftime(EndDate, "%d")),
            MounthCal=str(months[EndDate.month - 1]),
            YearCal=str(datetime.strftime(EndDate, "%Y")),
            Kleymo=str(LastKleymo),
            Name1=str(Name1),
            Name2=str(Name2),
            SerialNumber=str(self.TestWeightSet_SerialNumber),
            Method1=str(self.Method_ID),
            Etalon1=str(Etalon1),
            Etalon2=str(Etalon2),
            Temp=str(self.AirTemperature_Avr),
            Hym=str(self.Humidity_Avr),
            Press=str(self.AirPressure_Avr),
            HeadFIO=self.HeadFIO,
            HeadName=self.HeadName,
            UserFIO=str(self.CalibratedBy),
            Owner1=str(self.Company_Name),
            INN=str(self.CustomerNumber),
            ApprovedWith=str(ApprovedWith),
            Class=str(self.TestWeightSet_AccuracyClass)
        )

        rows = list()
        for i in self.TestWeights:
            rows.append({'MTNominal': str(i['NominalID']) + " " + str(i['NominalUnit']),
                         'MTConvertional': str(i['ConventionalMass']) + " " + str(i['ConventionalMassUnit']),
                         'MTError': str(i['ConventionalMassCorrection']) + " " + str(
                             i['ConventionalMassCorrectionUnit']),
                         'MTUncertainty': str(i['ExpandedMassUncertainty']) + " " + str(
                             i['ExpandedMassUncertaintyUnit'])})

        document.merge_rows('MTNominal', rows)

        date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
        file_to_save = self.rightFileName(
            self.Company_Name.strip() + ' ' + self.TestWeightSet_AccuracyClass.strip() + ' ' + self.TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.docx')

        fileApprovalCert = self.Excel_folder + '\\' + self.ApprovalCertFolder + r'\Свидетельство о поверке ' + file_to_save
        fileApprovalCert = fileApprovalCert.replace(',', ' ')

        document.write(fileApprovalCert)

        if self.autoopen == True:
            logging.debug('Файл свидетельства о поверке для ' + self.Company + ' создан: ' + fileApprovalCert)
            os.startfile(fileApprovalCert, 'open')

    # формирование поверочного протокола из docx шаблона
    def ReportDoc(self):
        try:
            logging.debug('формирование поверочного протокола для ' + self.Company)
            if 'Калибровка' in self.Method_Name:
                template = self.TemplateCalReport
            else:
                template = self.TemplateApprovalReport

            document = MailMerge(template)  # type: MailMerge

            MassComparators_Info = ''
            for m in self.MassComparators:
                MassComparators_Info += ', компаратор массы ' + m['Model'] + ' заводской номер ' + m['SerialNumber']

            RefrenceInfo = ''
            i = 0
            RefName = ''
            for r in self.ReferenceWeightSets:
                i += 1
                if i == 1:
                    RefName = str(r['ReferenceWeightSetName']).capitalize()
                else:
                    RefName = str(r['ReferenceWeightSetName'])
                RefrenceInfo += ' ' + RefName + ' ' + self.correctUnit(r['Range']) + ' заводской номер ' + r[
                    'SerialNumber'] + ' рег. № ' + r['Comment']
                if i < len(self.ReferenceWeightSets):
                    RefrenceInfo += '; '

            Result = ''
            Method = ''
            Type = ''
            DocName = ''
            if 'Калибровка' in str(self.Method_Name):
                Method = 'Калибровка'
                Type = 'калибровки'
                Period = ''
            else:
                Method = 'Поверка'
                Type = 'поверки'
                Period = 'периодической'

            if self.Test_Passed and Method == 'Поверка':
                Result = 'Заключение по результатам поверки: гири пригодны к использованию по классу точности ' + ClassReName(
                    self.TestWeightSet_AccuracyClass)
                DocName = 'Выдано свидетельство о поверке № ____________'
            if not self.Test_Passed and Method == 'Поверка':
                Result = 'Заключение по результатам поверки: гири непригодны к использованию по классу точности ' + ClassReName(
                    self.TestWeightSet_AccuracyClass)
                DocName = 'Выдано извещение о непригодности № ____________'
            if not self.Test_Passed and Method == 'Калибровка':
                Result = ''
                DocName = 'Выдан сертфикат о калибровке № ____________'

            DocNumber = self.ApprovalProtocolNum
            document.merge(
                Type=str(Type),
                Period=str(Period),
                Result=str(Result),
                Method=str(self.Method_ID),
                DocName=str(DocName),
                DocNumber=str(DocNumber),
                EndDate=str(self.EndDate),
                SerialNumber=str(self.TestWeightSet_SerialNumber),
                Company=str(self.Company_Name),
                Laboratory=str(self.CSM),
                TestClass=str(ClassReName(str(self.TestWeightSet_AccuracyClass))),
                ReferenceInfo=str(RefrenceInfo + MassComparators_Info),
                TempMin=str(self.AirTemperature_Min),
                TempMax=str(self.AirTemperature_Max),
                PressMin=str(self.AirPressure_Min),
                PressMax=str(self.AirPressure_Max),
                HymMin=str(self.Humidity_Min),
                HymMax=str(self.Humidity_Max),
                NominalUnit=str(self.TestWeight_NominalUnit[0]),
                MesurmentUnit=str(self.MesurmentUnit),
                DiffUnit=str(self.MesurmentUnit),
                RefUnit=str(self.MesurmentUnit),
                AvrUnit=str(self.MesurmentUnit),
                TestUnit=str(self.MesurmentUnit),
                TestWeightUnit=str(self.MesurmentUnit),
                UnsertUnit=str(self.MesurmentUnit),
                TolUnit=str(self.TolUnit),
                UserFIO=str(self.CalibratedBy)
            )

            rows = list()
            n = len(self.TestWeight_Nominal)
            i = 0
            while i < n:
                Mesurment = ''
                Diff = ''
                for r in range(0, len(self.A1[i])):
                    Diff += str(self.Diff[i][r]) + '\n' + '\n'
                    Mesurment += 'А1' + str(r + 1) + ' ' + str(self.A1[i][r]) + '\n'
                    Mesurment += 'B1' + str(r + 1) + ' ' + str(self.B1[i][r]) + '\n'
                    if len(self.B2[i]) > 0:
                        Diff += '\n'
                        Mesurment += 'B2' + str(r + 1) + ' ' + str(self.B2[i][r]) + '\n'
                    Mesurment += 'А2' + str(r + 1) + ' ' + str(self.A1[i][r])
                    if r < len(self.A1[i]) - 1:
                        Mesurment += '\n'
                        Diff += '\n'
                rows.append({'MTNominal': str(self.TestWeight_Nominal[i]) + str(self.TestWeight_NominalUnit[i]),
                             'MTMesurment': str(Mesurment).replace('.', ','),
                             'MTDiff': str(Diff).replace('.', '.'),
                             'MTAvr': str(self.Avr[i]).replace('.', ','),
                             'MTConventionalMassError': str(self.ReferenceWeight_ConventionalMassError[i]).replace('.',
                                                                                                                   ','),
                             'MTTestCorr': str(self.ConventionalMassCorrection[i]),
                             'MTConvMass': str(self.ConventionalMass[i]),
                             'MTUnsert': str(self.ExpandedMassUncertainty[i]),
                             'MTTol': str(self.TestWeight_Tolerance[i]).replace('.', ',')
                             })

                i += 1
            document.merge_rows('MTNominal', rows)
            date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
            file_to_save = self.rightFileName(self.Company_Name.strip() + ' ' + self.TestWeightSet_AccuracyClass.strip() + ' ' + self.TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.docx')

            fileApprovalProtocol = self.Excel_folder + '\\' + self.ApprovalProtocolFolder + '\\' + file_to_save
            fileApprovalProtocol = fileApprovalProtocol.replace(',', ' ')

            document.write(fileApprovalProtocol)
            if self.autoopen == True:
                logging.debug('Файл протокола поверки для ' + self.Company + ' создан: ' + fileApprovalProtocol)
                os.startfile(fileApprovalProtocol, 'open')
        except:
            logging.exception(
                'Ошибка формирования протокола для ' + self.Company + 'по шаблону ' + self.TemplateApprovalReport)

    # Калибровочный протокол
    def CalCertDoc(self):
        logging.debug('Формирование сертификата о калибровке для ' + self.Company)
        # указываем путь до шаблона
        template = self.TemplateCalCert

        EndDate = datetime.strptime(str(self.EndDate), "%d.%m.%Y").date()
        # создаем объект и смотрим на имеющиемся поля
        document = MailMerge(template)

        MassComparators_Info = ''
        for m in self.MassComparators:
            MassComparators_Info += ', компаратор массы ' + m['Model'] + ' заводской номер ' + m['SerialNumber']

        RefrenceInfo = ''
        i = 0
        RefName = ''
        for r in self.ReferenceWeightSets:
            i += 1
            if i == 1:
                RefName = str(r['ReferenceWeightSetName']).capitalize()
            else:
                RefName = str(r['ReferenceWeightSetName'])
            RefrenceInfo += ' ' + RefName + ' ' + self.correctUnit(r['Range']) + ' заводской номер ' + r[
                'SerialNumber'] + 'рег. №' + r['Comment']
            if i < len(self.ReferenceWeightSets):
                RefrenceInfo += '; '

        i = 1
        Etalon1 = ''
        Etalon2 = ''
        for r in self.ReferenceWeightSets:
            if i == 1: Etalon1 = r['Comment']
            if i == 2: Etalon1 += ', ' + r['Comment']
            if i == 3: Etalon2 = r['Comment']
            if i == 4: Etalon2 += ', ' + r['Comment']
            i += 1

        MaxLen1 = 45
        MaxLen2 = 63
        Name1 = ''
        Name2 = ''
        Name = self.TestWeightSet_Description.replace('\n', ' ')
        if len(Name) > MaxLen1:
            tmpstr = Name.split(' ')
            for i in tmpstr:
                if len(Name1 + i) < MaxLen1 and Name2 == '':
                    Name1 += i + ' '
                elif len(Name2 + i) < MaxLen2:
                    Name2 += i + ' '
        else:
            Name1 = Name

        months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', "июля", "августа", "сентября", "октября",
                  "ноября", "декабря"]

        UserName = 'поверитель'

        DocNumber = self.CalCertNum
        if 'Юшина' in self.CalibratedBy:
            UserName = 'инженер 1 категории'
        if 'Дидык' in self.CalibratedBy:
            UserName = 'ведущий инженер'

        document.merge(
            DocNumber=str(DocNumber),
            DocDate=str(datetime.strftime(EndDate, "%d.%m.%Y")),
            DayCal=str(datetime.strftime(EndDate, "%d")),
            MounthCal=str(months[EndDate.month - 1]),
            YearCal=str(datetime.strftime(EndDate, "%Y")),
            Name=str(Name),
            Name1=str(Name1),
            Name2=str(Name2),
            SerialNumber=str(self.TestWeightSet_SerialNumber),
            Customer=str(self.Company_Name),
            Method=str(self.Method_ID),
            Method1=str(self.Method_ID),
            Etalon=str(RefrenceInfo + MassComparators_Info),
            Etalon1=str(Etalon1),
            Etalon2=str(Etalon2),
            Temp=str(self.AirTemperature_Avr),
            Hum=str(self.Humidity_Avr),
            Press=str(self.AirPressure_Avr),
            HeadFIO=self.HeadFIO,
            HeadName=self.HeadName,
            UserName=str(UserName),
            UserFIO=str(self.CalibratedBy),
            Owner1=str(self.Company_Name),
            INN=str(self.CustomerNumber)
        )

        rows = list()
        for i in self.TestWeights:
            rows.append({'MTNominal': str(i['NominalID']) + " " + str(i['NominalUnit']),
                         'MTConvertional': str(i['ConventionalMass']) + " " + str(i['ConventionalMassUnit']),
                         'MTError': str(i['ConventionalMassCorrection']) + " " + str(
                             i['ConventionalMassCorrectionUnit']),
                         'MTUncertainty': str(i['ExpandedMassUncertainty']) + " " + str(
                             i['ExpandedMassUncertaintyUnit'])})

        document.merge_rows('MTNominal', rows)

        date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
        file_to_save = self.rightFileName(
            self.Company_Name.strip() + ' ' + self.TestWeightSet_AccuracyClass.strip() + ' ' + self.TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.docx')

        fileCalCert = self.Excel_folder + '\\' + self.CalCertFolder + '\\' + file_to_save
        fileCalCert = fileCalCert.replace(',', ' ')

        document.write(fileCalCert)

        if self.autoopen == True:
            logging.debug('Файл сертификата о калибровке для ' + self.Company + ' создан: ' + fileCalCert)
            os.startfile(fileCalCert, 'open')

    def ErrorReportDoc(self):
        logging.debug('Формирование извещения о непригодности для' + self.Company)
        # указываем путь до шаблона
        template = self.TemplateError

        if self.TestWeightSet_InternalID == "":
            LastKleymo = '-'
        else:
            LastKleymo = self.TestWeightSet_InternalID

        MaxLen1 = 45
        MaxLen2 = 63
        Name1 = ''
        Name2 = ''
        Name = self.TestWeightSet_Description.replace('\n', ' ')
        if len(Name) > MaxLen1:
            tmpstr = Name.split(' ')
            if tmpstr[0] == 'Поверка' or tmpstr[0] == 'Калибровка':
                tmpstr = tmpstr[1:]
            for i in tmpstr:
                if len(Name1 + i) < MaxLen1 and Name2 == '':
                    Name1 += i + ' '
                elif len(Name2 + i) < MaxLen2:
                    Name2 += i + ' '
        else:
            Name1 = Name

        MaxLen1 = 37
        MaxLen2 = 63
        Method1 = ''
        Method2 = ''
        Method = self.Method_ID
        if len(Method) > MaxLen1:
            tmpstr = Method.split(' ')
            if tmpstr[0] == 'Поверка' or tmpstr[0] == 'Калибровка':
                tmpstr = tmpstr[1:]

            for i in tmpstr:
                if len(Method1 + i) < MaxLen1 and Method2 == '':
                    Method1 += i + ' '
                elif len(Method2 + i) < MaxLen2:
                    Method2 += i + ' '
        else:
            Method1 = Method

        EndDate = datetime.strptime(str(self.EndDate), "%d.%m.%Y").date()
        # создаем объект и смотрим на имеющиемся поля
        document = MailMerge(template)
        # print(document.get_merge_fields())

        months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', "июля", "августа", "сентября", "октября",
                  "ноября", "декабря"]

        Reason1 = 'Превышение допустимой погрешности'
        Reason2 = ''
        for e in self.ErrorResults:
            Reason2 += e['NominalID'] + 'г ' + '\u03B4m ' + e['Error']

        DocNumber = self.ErrorNum
        document.merge(
            DocNumber=str(DocNumber),
            DayCal=str(datetime.strftime(EndDate, "%d")),
            MounthCal=str(months[EndDate.month - 1]),
            YearCal=str(datetime.strftime(EndDate, "%Y")),
            Kleymo=str(LastKleymo),
            Name1=str(Name1),
            Name2=str(Name2),
            SerialNumber=str(self.TestWeightSet_SerialNumber),
            Method1=str(Method1),
            Method2=str(Method2),
            Reason1=str(Reason1),
            Reason2=str(Reason2),
            HeadFIO=self.HeadFIO,
            HeadName=self.HeadName,
            UserFIO=str(self.CalibratedBy),
            Owner=str(self.Company_Name),
            INN=str(self.CustomerNumber),
            ApprovedWith='в полном объеме'
        )

        date_time = datetime.now().strftime("%d%m%Y_%H%M%S")
        file_to_save = self.rightFileName(
            self.Company_Name.strip() + ' ' + self.TestWeightSet_AccuracyClass.strip() + ' ' + self.TestWeightSet_SerialNumber.strip() + ' ' + date_time + '.docx')

        fileError = self.Excel_folder + '\\' + self.ErrorFolder + '\\' + file_to_save
        fileError = fileError.replace(',', ' ')

        document.write(fileError)

        if self.autoopen == True:
            logging.debug('Файл извещения о непригодности для ' + self.Company + ' создан: ' + fileError)
            os.startfile(fileError, 'open')

    # формирование протокола поверки в формате xlsx
    def ApprovalProtocolXlsx(self):
        wb = openpyxl.Workbook
        wb = load_workbook(filename=r'C:\Doc\prog\MCLinkReport\templates\Протокол поверки_клин.xlsx', read_only=False)

        self.Set_Cell(wb, 'DocNum', 'B00123/12380823')
        self.Set_Cell(wb, 'EndDate', '12.12.2012')
        self.Set_Cell(wb, 'ReestrNum', '12343-18')
        self.Set_Cell(wb, 'CustomerName', 'ФБУ Клинский ЦСМ')
        self.Set_Cell(wb, 'Range', '1г - 500г')
        self.Set_Cell(wb, 'SerialNumber', '121234234')
        self.Set_Cell(wb, 'Class', 'F1')
        self.Set_Cell(wb, 'TempAvr', '21,5')
        self.Set_Cell(wb, 'HymAvr', '40')
        self.Set_Cell(wb, 'PressAvr', '991')
        self.Set_Cell(wb, 'DensityAvr', '1,15342')
        self.Set_Cell(wb, 'Method', 'МП 17002')
        self.Set_Cell(wb, 'EtalonInfo', '2.1.ZТТ.1956.2017')
        self.Set_Cell(wb, 'Cell1', '123,1234')

    # запуск конвертации
    def convertation(self, xml_filename=None):
        try:
            logging.debug('Запуск раcпознавания файла ' + xml_filename)
            error = self.ParseXML(xml_filename)
            if error != None:
                logging.debug(error)
                os.remove(xml_filename)
                shutil.copy(logfileName, self.xml_folder + r'\error.txt')
            else:
                logging.debug('Распознавание файла ' + xml_filename + ' завершено')
                if 'Поверка компаратора' in self.Method_Name:
                    logging.debug('Запуск создания протокола поверки компаратора')
                    self.ComparatorApprovalReport(xml_filename)
                    logging.debug('Завершение создания протокола поверки компаратора')
                elif 'Калибровка' in str(self.Method_Name):
                    if 'Клинский' in self.CSM:
                        self.CalReportKlin(xml_filename)
                    else:
                        # self.CallReport(xml_filename)
                        self.CalCertDoc()
                        self.ReportDoc()
                else:
                    if 'Клинский' in self.CSM:
                        self.ApprovalReportKlin(xml_filename)
                    else:
                        if '.xls' in self.TemplateApprovalCert:
                            self.ApprovalReport(xml_filename)
                        elif '.docx' in self.TemplateApprovalCert and self.Test_Passed:
                            self.ApprovalCertDoc()
                            self.ReportDoc()

                        elif '.docx' in self.TemplateApprovalCert and not self.Test_Passed:
                            self.ErrorReportDoc()
                            self.CalCertDoc()
                            self.ReportDoc()

                logging.debug('Удаление файла xml ' + xml_filename)
                os.remove(xml_filename)
        except:
            logging.exception('Ошибка преобразования файла ' + xml_filename)
            # Удаляем исходный файл xml
            # if self.autodelXML == True:
            logging.debug('Удаление файла xml ' + xml_filename)
            os.remove(xml_filename)
            try:
                os.replace(self.logfile, self.xml_folder + r'\error.txt')
            except:
                pass

    # парсинг файла XML
    def ParseXML(self, xml_filename):
        self.MassComparators.clear()
        self.ReferenceWeightSets.clear()
        self.ReferenceWeighs.clear()
        self.TestWeight_Nominal.clear()
        self.TestWeight_NominalUnit.clear()
        self.TestWeight_Tolerance.clear()
        self.TestWeight_WeightId.clear()
        self.ConventionalMassCorrection.clear()
        self.ConventionalMass.clear()
        self.ConventionalMassUnit.clear()
        self.ConventionalMassCorrectionUnit.clear()
        self.A1.clear()
        self.A2.clear()
        self.B1.clear()
        self.B2.clear()
        self.Diff.clear()
        self.Avr.clear()
        self.TestWeights.clear()
        self.ReferenceWeight_ConventionalMassError.clear()
        self.ReferenceWeight_ConventionalMassErrorUnit.clear()
        self.ExpandedMassUncertainty.clear()
        self.ExpandedMassUncertaintyUnit.clear()

        try:
            logging.debug('Попытка открыть файл ' + xml_filename)
            tree = etree.parse(xml_filename)
            logging.debug('Файл открыт: ' + xml_filename)
        except:
            logging.exception('Ошибка открытия файла ' + xml_filename)
            sleep(1)
            logging.debug('Повторная попытка открыть файл ' + xml_filename)
            try:
                tree = etree.parse(xml_filename)
            except:
                logging.exception('Повторная ошибка открытия файла ' + xml_filename)
                return 'Ошибка открытия файла'

        try:
            root = tree.getroot()
            WeightSetCalibration = root.find('WeightSetCalibration')
            ContactOwner = WeightSetCalibration.find('ContactOwner')
            Company = str(ContactOwner.find('Company').text).strip(' ')
            City = ContactOwner.find('City')

            self.Company = Company
            self.CustomerNumber = str(ContactOwner.find('Company').text).strip(' ')
            self.CustomerNumber = ContactOwner.get('CustomerNumber')  # ИНН
            TestWeightSet = WeightSetCalibration.find('TestWeightSet')

            self.TestWeightSet_Description = TestWeightSet.find('Description').text  # название набора с номером в гр
            self.TestWeightSet_Comment = TestWeightSet.find('Comment').text  # номер эталона
            self.TestWeightSet_SerialNumber = TestWeightSet.get('SerialNumber')
            self.TestWeightSet_AccuracyClass = TestWeightSet.get('AccuracyClass').split(' ')[2]
            self.TestWeightSet_Manufacturer = TestWeightSet.get('Manufacturer')
            self.TestWeightSet_InternalID = TestWeightSet.get('InternalID')  # номер клейма предыдущей поверки
            self.TestWeightSet_Range = self.correctUnit(TestWeightSet.get('Range'))

            TestWeightCalibrations = TestWeightSet.find('TestWeightCalibrations')
            TestWeightCalibrationAsReturned = TestWeightCalibrations.findall('TestWeightCalibrationAsReturned')
            TestWeightCalibrationAsFound = TestWeightCalibrations.findall('TestWeightCalibrationAsFound')
            TestWeightSet_AlloyMaterials = TestWeightSet.find('AlloyMaterials')
            TestWeightSet_AlloyMaterial = TestWeightSet_AlloyMaterials.findall('AlloyMaterial')[0]
            self.Density = TestWeightSet_AlloyMaterial.get('Density') + \
                           TestWeightSet_AlloyMaterial.get('DensityUnit')  # плотность испытуемых гирь

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

            self.TestWeight_CalibrationsCount = TestWeightCalibrations.get('Count')  # количество гирь
            self.EndDate = WeightSetCalibration.get('EndDate')  # дата поверки
            self.CertificateNumber = WeightSetCalibration.get('CertificateNumber')  # номер сертификата
            self.CalibratedBy = WeightSetCalibration.get('CalibratedBy')  # поверитель
            self.CustomerNumber = ContactOwner.get('CustomerNumber')  # ИНН

            self.Company_Name = Company  # назвение заказчика

            self.Address = City.get('ZipCode') + ', ' + City.get('State') + ', ' + ContactOwner.find(
                'Address').text  # адрес
            self.AirDensity_Min = self.roundStr(AirDensity.find('Min').text, 4)
            self.AirDensity_Max = self.roundStr(AirDensity.find('Max').text, 4)
            self.AirDensity_Avr = self.roundStr(AirDensity.find('Average').text, 4)
            self.AirDensity_Unit = AirDensity.find('Unit').text
            self.AirTemperature_Min = self.roundStr(AirTemperature.get('Min'), 2)  # температура мин
            self.AirTemperature_Max = self.roundStr(AirTemperature.get('Max'), 2)  # температура макс
            self.AirTemperature_Avr = self.roundStr(AirTemperature.get('Average'), 2)  # температура средняя
            self.AirTemperature_Unit = AirTemperature.get('Unit')  # размерность температуры

            self.AirPressure_Min = self.roundStr(AirPressure.get('Min'), 2)  # давление мин
            self.AirPressure_Max = self.roundStr(AirPressure.get('Max'), 2)  # давление макс
            self.AirPressure_Avr = self.roundStr(AirPressure.get('Average'), 2)  # давление среднее
            self.AirPressure_Unit = AirPressure.get('Unit')  # размерность давления

            self.Humidity_Min = self.roundStr(Humidity.get('Min'), 2)  # влажность мин
            self.Humidity_Max = self.roundStr(Humidity.get('Max'), 2)  # влажность макс
            self.Humidity_Avr = self.roundStr(Humidity.get('Average'), 2)  # влажность средняя
            self.Humidity_Unit = Humidity.get('Unit')  # размерность влажности

            self.Method_Name = Method[0].get('Name')  # метод поверки
            self.Method_Process = Method[0].get('Process')  # название процесса поверки
            self.Method_ID = Method[0].text  # Признак метода Калибровка или Поверка и название метода через пробел
        except:
            logging.exception('Ошибка чтения параметров калибровки')
            return 'Ошибка чтения параметров калибровки'

        try:
            for comp in MassComparator:
                self.MassComparators.append({'Index': comp.get('Index'),
                                             'Model': comp.get('Model'),
                                             'SerialNumber': comp.get('SerialNumber'),
                                             'Description': str(comp.find('Description').text)})

            for comp in MassComparator:
                self.MassComparators_Info += str(comp.get('Model')) + str(comp.get('SerialNumber')) + str(
                    comp.find('Description').text)
                self.MassComparator_Model.append(comp.get('Model'))  # модели компараторов
                self.MassComparator_SerialNumber.append(comp.get('SerialNumber'))  # серийные номера компараторов
                self.MassComparator_Description.append(
                    comp.find('Description').text)  # описание компаратора (дискретность, ско...)
        except:
            logging.exception('Ошибка чтения параметров компараторов')
            return 'Ошибка чтения параметров компараторов'

        try:

            for ref in ReferenceWeight:
                self.ReferenceWeighs.append({'Index': ref.get('Index'),
                                             'SerialNumber': ref.get('SerialNumber'),
                                             'NominalWeight': ref.get('NominalWeight'),
                                             'NominalWeightUnit': ref.find('NominalWeightUnit'),
                                             'WeightId': ref.find('WeightId'),
                                             'Density': ref.find('Density'),
                                             'Class': ref.find('Class'),
                                             'ConventionalMass': ref.find('ConventionalMass'),
                                             'ConventionalMassUnit': ref.find('ConventionalMassUnit'),
                                             'ConventionalMassError': ref.find('ConventionalMassError'),
                                             'ConventionalMassErrorUnit': ref.find('ConventionalMassErrorUnit'),
                                             'ExpandedMassErrorUncertainty': ref.find('ExpandedMassErrorUncertainty'),
                                             'ExpandedMassErrorUncertaintyUnit': ref.find(
                                                 'ExpandedMassErrorUncertaintyUnit'),
                                             'CertificateNumber': ref.find('CertificateNumber')
                                             })

            for ref in ReferenceWeightSet:
                refrange = str(ref.get('Range')).split('-')
                ReferenceWeightSetName = ''
                if len(refrange) == 2:
                    ReferenceWeightSetName = 'набор гирь'
                else:
                    ReferenceWeightSetName = 'гиря'

                self.ReferenceWeightSets.append({'SerialNumber': ref.get('SerialNumber'),
                                                 'Clаss': ref.get('Class'),
                                                 'Range': ref.get('Range'),
                                                 'CommonAlloyMaterial': ref.get('CommonAlloyMaterial'),
                                                 'CommonAlloyMaterialDensity': ref.get('CommonAlloyMaterialDensity'),
                                                 'CommonAlloyMaterialDensityUnit': ref.get(
                                                     'CommonAlloyMaterialDensityUnit'),
                                                 'CommonShape': ref.get('CommonShape'),
                                                 'LastCalibrationDate': ref.get('LastCalibrationDate'),
                                                 'CertificateNumber': ref.get('CertificateNumber'),
                                                 'NextCalibrationDate': ref.get('NextCalibrationDate'),
                                                 'Comment': ref.find('Comment').text,
                                                 'ReferenceWeightSetName': ReferenceWeightSetName
                                                 })
                self.ReferenceWeightSet_SerialNumber.append(ref.get('SerialNumber'))  # серийный номер набора эталонов
                self.ReferenceWeightSet_Class.append(ref.get('Class'))  # классы эталонных наборов
                self.ReferenceWeightSet_Range.append(ref.get('Range'))  # диапазон эталонных гирь
                self.ReferenceWeightSet_RegNumber.append(ref.find('Comment').text)
                self.ReferenceWeightSet_NextCalibrationDate.append(ref.get('NextCalibrationDate'))

        except:
            logging.exception('Ошибка чтения параметров эталонов')
            return 'Ошибка чтения параметров эталонов'

        self.Test_Passed = True

        try:
            self.ErrorResults = []
            # Есть отрицательные результаты или ошибочно записанные AsFound
            if len(TestWeightCalibrationAsFound) > 0:
                for found in TestWeightCalibrationAsFound:
                    Nominal = float(str(found.find('Nominal').text).replace(',', '.'))
                    u = 1
                    eu = 1
                    NominalUnit = str(found.find('NominalUnit').text)

                    if NominalUnit == 'kg': u = 1000
                    if NominalUnit == 'mg': u = 0.001
                    if NominalUnit == 'ug': u = 0.000001
                    Error = float(str(found.find('ConventionalMassCorrection').text).replace(',', '.'))
                    ErrorUnit = self.correctUnit(found.find('ConventionalMassCorrectionUnit').text)
                    if ErrorUnit == 'kg': eu = 1000
                    if ErrorUnit == 'mg': eu = 0.001
                    if ErrorUnit == 'ug': eu = 0.000001

                    Tolerance = float(str(found.find('Tolerance').text).replace(',', '.'))
                    # Отрицательный результат
                    if abs(Error * eu) < 0.1 * Nominal * u and abs(Error) > Tolerance:
                        TestWeightCalibrationAsReturned.append(found)
                        TestWeightCalibrationAsFound.remove(found)
                        self.Test_Passed = False

                        self.ErrorResults.append({'NominalID': str(found.find('WeightID').text).strip('\n').strip(' ') +
                                                               str(Nominal).replace('.', ','),
                                                  'Error': str(Error).replace('.', ',') + str(ErrorUnit)})
                    # Ошибочно записанный положительный результат
                    elif abs(Error) <= Tolerance:
                        TestWeightCalibrationAsReturned.append(found)
                        # TestWeightCalibrationAsFound.remove(found)
                    else:
                        # TestWeightCalibrationAsFound.remove(found)
                        self.Test_Passed = False
            else:
                self.Test_Passed = True
        except:
            logging.exception('Ошибка проверки результатов')
            return 'Ошибка проверки результатов'

        # Заполняем протокол поверки если есть положительные результаты AsReturned
        if len(TestWeightCalibrationAsReturned) > 0:
            try:
                if self.TestWeight_CalibrationsCount == '1':
                    self.CI_Name = 'Гиря'
                else:
                    self.CI_Name = 'Набор гирь'

                for i in TestWeightCalibrationAsReturned:
                    Nominal = i.find('Nominal').text
                    WeightID = i.find('WeightID').text
                    self.TestWeight_WeightId.append(WeightID)
                    Index = i.find('ReferenceWeight').text
                    Tolerance = i.find('Tolerance').text
                    ReferenceWeight_ConventionalMassError = 0
                    ReferenceWeight_ConventionalMassErrorUnit = ''
                    for j in ReferenceWeight:
                        if Index == j.get('Index'):
                            ReferenceWeight_ConventionalMassError = j.get('ConventionalMassError')
                            self.ReferenceWeight_ConventionalMassError.append(ReferenceWeight_ConventionalMassError)
                            ReferenceWeight_ConventionalMassErrorUnit = self.correctUnit(
                                j.get('ConventionalMassErrorUnit'))
                            self.ReferenceWeight_ConventionalMassErrorUnit.append(
                                ReferenceWeight_ConventionalMassErrorUnit)
                    if self.TestWeight_CalibrationsCount == '1':
                        self.TestWeight_Nominal.append((str(Nominal).strip('\n').strip(' ')))
                    else:
                        self.TestWeight_Nominal.append(str(WeightID + Nominal).strip('\n').strip(' '))
                    self.TestWeight_NominalUnit.append(self.correctUnit(i.find('NominalUnit').text))

                    self.TestWeight_Tolerance.append(float(Tolerance.replace(',', '.')))
                    MeasurementReadings = i.find('MeasurementReadings')
                    try:
                        WeightReading = MeasurementReadings.findall('WeightReading')
                        A1 = []
                        A2 = []
                        B1 = []
                        B2 = []
                        Diff = []
                        round_number = 5
                        # Определение метода
                        StepSeriesIndex = ''
                        for wr in WeightReading:
                            StepSeriesIndex += wr.get('Step') + wr.get('SeriesIndex')

                        Method = StepSeriesIndex[0:6]
                        # ABA
                        if Method == 'A1B1A1' or Method == '(A)1B1':  # 1 ABA
                            for cicle in range(int(len(WeightReading) / 3)):
                                round_number = abs(decimal.Decimal(
                                    str(WeightReading[cicle * 3].get('WeightReading')).replace(',',
                                                                                               '.')).as_tuple().exponent) + 1
                                A1.append(float(str(WeightReading[cicle * 3].get('WeightReading')).replace(',', '.')))
                                B1.append(
                                    float(str(WeightReading[cicle * 3 + 1].get('WeightReading')).replace(',', '.')))
                                A2.append(
                                    float(str(WeightReading[cicle * 3 + 2].get('WeightReading')).replace(',', '.')))
                                Diff.append(round((B1[cicle] - (A1[cicle] + A2[cicle]) / 2), round_number))

                        if Method == 'A1B1B1':  # 1 ABBA
                            for cicle in range(int(len(WeightReading) / 4)):
                                round_number = abs(decimal.Decimal(
                                    str(WeightReading[cicle * 4].get('WeightReading')).replace(',',
                                                                                               '.')).as_tuple().exponent) + 1
                                A1.append(float(str(WeightReading[cicle * 4].get('WeightReading')).replace(',', '.')))
                                B1.append(
                                    float(str(WeightReading[cicle * 4 + 1].get('WeightReading')).replace(',', '.')))
                                B2.append(
                                    float(str(WeightReading[cicle * 4 + 2].get('WeightReading')).replace(',', '.')))
                                A2.append(
                                    float(str(WeightReading[cicle * 4 + 3].get('WeightReading')).replace(',', '.')))
                                Diff.append(
                                    round(((B1[cicle] + B2[cicle]) / 2 - (A1[cicle] + A2[cicle]) / 2), round_number))
                        self.A1.append(A1)
                        self.A2.append(A2)
                        self.B1.append(B1)
                        self.B2.append(B2)
                        self.Diff.append(Diff)
                        Avr = round(mean(Diff), round_number)
                        self.Avr.append(Avr)

                        ConventionalMassCorrection = i.find('ConventionalMassCorrection').text
                        self.ConventionalMassCorrection.append(self.roundStr(ConventionalMassCorrection, round_number))
                        ConventionalMassCorrectionUnit = i.find('ConventionalMassCorrectionUnit').text
                        self.ConventionalMassCorrectionUnit.append(self.correctUnit(ConventionalMassCorrectionUnit))
                        ConventionalMass = i.find('ConventionalMass').text
                        self.ConventionalMass.append(ConventionalMass)
                        ConventionalMassUnit = i.find('ConventionalMassUnit').text
                        self.ConventionalMassUnit.append(self.correctUnit(ConventionalMassUnit))

                        self.ExpandedMassUncertainty.append(i.find('ExpandedMassUncertainty').text)
                        ExpandedMassUncertaintyUnit = self.correctUnit(i.find('ExpandedMassUncertaintyUnit').text)
                        self.ExpandedMassUncertaintyUnit.append(ExpandedMassUncertaintyUnit)
                        self.TestWeights.append({
                            'Nominal': str(i.find('Nominal').text).strip('\n').strip(' '),
                            'NominalUnit': self.correctUnit(str(i.find('NominalUnit').text)),
                            'WeightID': str(i.find('WeightID').text).strip('\n').strip(' '),
                            'NominalID': str(i.find('WeightID').text).strip('\n').strip(' ') + str(
                                i.find('Nominal').text).strip('\n').strip(' '),
                            'Density': str(i.find('Density').text),
                            'DensityUnit': str(i.find('DensityUnit').text),
                            'MassComparator': str(i.find('MassComparator').text),
                            'ReferenceWeight': str(i.find('ReferenceWeight').text),
                            'ConventionalMassError': str(ReferenceWeight_ConventionalMassError),
                            'ConventionalMassErrorUnit': self.correctUnit(
                                str(ReferenceWeight_ConventionalMassErrorUnit)),
                            'ConventionalMassCorrection': str(i.find('ConventionalMassCorrection').text),
                            'ConventionalMassCorrectionUnit': self.correctUnit(
                                str(i.find('ConventionalMassCorrectionUnit').text)),
                            'ConventionalMass': str(i.find('ConventionalMass').text),
                            'ConventionalMassUnit': self.correctUnit(str(i.find('ConventionalMassUnit').text)),
                            'CombinedMassUncertainty': str(i.find('CombinedMassUncertainty').text),
                            'CombinedMassUncertaintyUnit': self.correctUnit(
                                str(i.find('CombinedMassUncertaintyUnit').text)),
                            'ExpandedMassUncertainty': str(i.find('ExpandedMassUncertainty').text),
                            'ExpandedMassUncertaintyUnit': self.correctUnit(
                                str(i.find('ExpandedMassUncertaintyUnit').text)),
                            'ExpansionFactor': str(i.find('ExpansionFactor').text),
                            'Tolerance': str(i.find('Tolerance').text).replace(',', '.'),
                            'ToleranceUnit': self.correctUnit(str(i.find('ToleranceUnit').text)),
                            'CalibrationResult': str(i.find('CalibrationResult').text),
                            'A1': A1, 'B1': B1, 'B2': B2, 'A2': A2, 'D': Diff, 'Avr': Avr
                        })
                    except:
                        logging.exception('Ошибка расчета результатов ' + xml_filename)
                        return 'Ошибка расчета результатов'
            except:
                logging.exception('Ошибка чтения параметров калибровки')
                return 'Ошибка чтения параметров калибровки'

    # установка параметров ячейки
    def Set_Cell(self, _wb, name, value):
        range = _wb.defined_names[name]
        # if this contains a range of cells then the destinations attribute is not None
        dests = range.destinations  # returns a generator of (worksheet title, cell range) tuples
        for title, coord in dests:
            _ws = _wb[title]
        _ws[coord].value = value


# класс главного окна
class MainWindow(QMainWindow):
    demon = DemonConvertation()
    Title = " Сохранение отчетов MCLink v" + ver

    # конструктор главного окна
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

        self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalReport)
        self.ui.lbTemplateApprovalCert.setText(self.demon.TemplateApprovalCert)
        self.ui.lbTemplateCalProtocol.setText(self.demon.TemplateCalReport)
        self.ui.lbTemplateCalCert.setText(self.demon.TemplateCalCert)
        self.ui.lbTemplateError.setText(self.demon.TemplateError)

        self.ui.lbScanFolder.setText(self.demon.xml_folder)
        self.ui.lbDestFolder.setText(self.demon.Excel_folder)
        self.ui.lbTemplateApprovalProtocol.setText(self.demon.template_filename)
        self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalReport)
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

    # устанавливаем признаки протоколов согласно установленным галкам на окне
    def setProtocol(self):
        self.demon.setProtocol(1, self.ui.cbxApprovalProtocol.isChecked())
        self.demon.setProtocol(2, self.ui.cbxApprovalSert.isChecked())
        self.demon.setProtocol(3, self.ui.cbxError.isChecked())
        self.demon.setProtocol(4, self.ui.cbxCalProtocol.isChecked())
        self.demon.setProtocol(5, self.ui.cbxCalSert.isChecked())

    # сохраняем настройки из окна
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

    # запускаем процесс автоматической конвертации
    def start(self):
        self.ui.startAction.setVisible(False)
        self.ui.stopAction.setVisible(True)
        self.demon = DemonConvertation()
        self.demon.setAutoStart(True)
        self.demon.setDaemon(True)
        self.demon.start()
        logging.debug('Запуск формирования протоколов')

    # останавливаем процесс автоматической конвертации
    def stop(self):
        self.ui.startAction.setVisible(True)
        self.ui.stopAction.setVisible(False)
        self.demon.setAutoStart(False)
        logging.debug('Остановка формирования протоколов')

    # выбираем папку XML с входными фалами
    def selectXmlFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите исходный файл', '')
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setXmlFolder(folder)
            self.ui.lbScanFolder.setText(self.demon.xml_folder)
            logging.debug('Изменена папка xml: ' + self.demon.xml_folder)

    # выбираем папку Excel с выходными файлами
    def selectExcelFolder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выберите папку сохранения отчетов', '')
        if folder != '':
            folder = str(folder).replace('/', '\\')
            self.demon.setExcelFolder(folder)
            self.ui.lbDestFolder.setText(self.demon.Excel_folder)
            logging.debug('Изменена папка протоколов: ' + self.demon.Excel_folder)

    # выбираем шаблон протокола поверки
    def selectTemplateApprovalProtocol(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона потокола поверки',
                                                    self.demon.TemplateApprovalReport, '*.xls;*.docx')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateApprovalProtocol')
            self.ui.lbTemplateApprovalProtocol.setText(self.demon.TemplateApprovalReport)
            logging.debug('Изменен шаблон протокола поверки: ' + self.demon.TemplateApprovalReport)

    # выбираем шаблон свидетельства о поверке
    def selectTemplateApprovalCert(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона свидетельства о поверке',
                                                    self.demon.TemplateApprovalCert, '*.xls;*.docx')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateApprovalCert')
            self.ui.lbTemplateApprovalCert.setText(self.demon.TemplateApprovalCert)
            logging.debug('Изменен шаблон свидетельства о поверке: ' + self.demon.TemplateApprovalCert)

    # выбираем шаблон извещения о неригодности
    def selectTemplateError(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона извещения о непригодности',
                                                    self.demon.TemplateError, '*.xls;*.docx')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateError')
            self.ui.lbTemplateError.setText(self.demon.TemplateError)

    # вибираем шаблон протокола калибровки
    def selectTemplateCalProtocol(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона потокола калибровки',
                                                    self.demon.TemplateCalReport, '*.xls;*.docx')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateCalProtocol')
            self.ui.lbTemplateCalProtocol.setText(self.demon.TemplateCalReport)

    # вибираем шаблон сертификата калибровки
    def selectTemplateCalCert(self):
        template, ext = QFileDialog.getOpenFileName(self, 'Выберите файл шаблона сертификата калибровки',
                                                    self.demon.TemplateCalCert, '*.xls;*.docx')
        if template != '':
            template = str(template).replace('/', '\\')
            self.demon.setTemplateFilename(template, 'TemplateCalCert')
            self.ui.lbTemplateCalCert.setText(self.demon.TemplateCalCert)

    # меняем признак автооткрытия
    def changeAutoOpen(self):
        if self.ui.chbAutoOpen.isChecked():
            self.demon.setAutoOpen(True)
            logging.debug('Автооткрытие включено')
        else:
            self.demon.setAutoOpen(False)
            logging.debug('Автооткрытие отключено')


# функци запуска приложения
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()

    logging.debug('======================================')
    logging.debug(ex.demon.CSM)
    logging.debug('Запуск программы версии ' + ver)
    ex.show()
    sys.exit(app.exec_())
