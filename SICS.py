import serial
import time
import datetime
from threading import Thread


class SICS(Thread):
    virtualName = 'COM7'
    serialName = 'COM3'
    virtual = serial.Serial()
    scale = serial.Serial()
    Connected = False
    Commands1Line = \
        ['I1', 'I2', 'I3', 'I11', 'I28', 'I51',
         'M01', 'M02', 'M03', 'M07', 'M12', 'M21', 'M31', 'M14',
         'DW', 'D',
         'SI', 'SIR', 'SIU',
         'UPD', 'K', '@',
         'XM0000', 'XM1070', 'XM1001', 'XM1008', 'XM2100', 'XM2101', 'XM2102', 'XM0021', 'XM0007', 'XM1001', 'XM1002', 'XM1000', 'XM1003', 'XM1014',
         'XP0000', 'XP0314', 'XP0126', 'XP0700', 'XP0366', 'XP0336', 'XP0306', 'XP0361', 'XP0324', 'XP2300', 'XP0323', 'XP0344', 'XP0345', 'XP0600',
         'XP0301', 'XP0303', 'XP0304', 'XP0305', 'XP0307', 'XP0323', 'XP0346', 'XP0201', 'XP0223', 'XP0243', 'XP0219', 'XP0236', 'XP0216', 'XP0217', 'XP0200',
         'XP0255', 'XA0001', 'XA1005',
         'ZZ13', 'COPT', 'XD11', 'WS']
    CommandsNLine = ['I0', 'ZZ00', 'ZZ01', 'XP1700', 'XU1008', 'XU0000', 'XM0000', 'XU0002', 'XU0003', 'XU1008','XA0000','XP0220','XP0240','XP0235','XP0207', 'XP0208','XP0209','XP0234','XP0200']
    CommandTemp = []
    AnswersFromScale = ['S']
    ZZ = [
        "ZZ00",  #: ZZ - commands LOCKER 4799 активация
        "ZZ01",  #: System - Info < 0:Process | | 1: Performance | | 2:Memory | | 3: EAROM | | 4:CSData >
        "ZZ02",  #: Set SP - Dumpmode(paramdescr. -> psi)
        "ZZ03",  #: NVStorage HEX - view < StorageID(1 = CELL / 2 = TYPE) >
        "ZZ04",  #: NVStorage analyzer < StorageID(1 = CELL / 2 = TYPE) >
        "ZZ05",  #: FACT - Settings < 1:Weekdays < bitselection >> | | < 2: Daytime < index, hour, min >> | | < 3: Temp < tempdiff in 1 / 10    Grad >>
        "ZZ06",  #: CS - Data monitor < id > < mode(0 = remove / 1 = add / 2 = add BIN / 3 = add  ASCII) >
        "ZZ07",  #: Touch driver: < Cmd(1 = Report / 2 Cal) > < Param( for Report: 1 = On / 0 = Off; for Cal: 1 = Prod / 2 = Usr) >
        "ZZ08",  #: Touch -, Key - Simulation < what > < pos > < pos >
        "ZZ09",  #: Set PP - Monitormode < Mode(0..6) >
        "ZZ10",  #: Send SP - Output < Mode(0..4) >
        "ZZ11",  #: Flash - Check < 0:Verify > | | < 1: Build and Store >
        "ZZ12",  #: I2C - Access < 0: rd, 1: wr > < adr > < siz > < dat0 >..< dat7 >
        "ZZ97",  #: Available for < DEBUG >
        "ZZ98",  #: Available for < DEBUG >
        "ZZ99",  #: Free ParamParsing < numOfParam(0 = gen) > < paramFormStr > < params... >
    ]


    XU = [
        "XU0000",  #4715 акстивация
        "XU0002",
        "XU0003",
        "XU0004",
        "XU0005",
        "XU0006",
        "XU0007",
        "XU0008",
        "XU0009",
        "XU0010",
        "XU0011",
        "XU0012",
        "XU1000",
        "XU1008"
    ]

    XA = [
        "XA0000",  #4714 активация
        "XA0001",
        "XA0002",  #сброс настроек
        "XA0003",
        "XA0004",
        "XA1000",
        "XA1001",
        "XA1002",
        "XA1003",
        "XA1004",
        "XA1005",
        "XA1006",
        "XA1007"
    ]

    XM = [
        "XM0000",  #4713 активация
        "XM0001",
        "XM0002",
        "XM0003",
        "XM0004",
        "XM0005",
        "XM0006",
        "XM0007",
        "XM0008",
        "XM0009",
        "XM0010",
        "XM0011",
        "XM0012",
        "XM0013",
        "XM0014",
        "XM0021",  #
        "XM0015",
        "XM1070",
        "XM1000",
        "XM1001",
        "XM1002",
        "XM1003",
        "XM1004",
        "XM1005",
        "XM1006",
        "XM1007",
        "XM1008",
        "XM1009",
        "XM1011",
        "XM1012",  #
        "XM1014",  #
        "XM2000",
        "XM2001",
        "XM2002",
        "XM2003",
        "XM2004",
        "XM2005",
        "XM2006",
        "XM2007",
        "XM2008",
        "XM2100",
        "XM2101",
        "XM2102",
        "XM2103",
        "XM2104",
        "XM2105",
        "XM2106"
    ]


    XP = [
        "XP0000", #4712 активация допкоманд
        "XP0101",
        "XP0102",
        "XP0103",
        "XP0107",
        "XP0108",
        "XP0109",
        "XP0113",
        "XP0116",
        "XP0120",
        "XP0122",
        "XP0126", #температура1
        "XP0127",
        "XP0134",
        "XP0135",
        "XP0138",
        "XP0139",
        "XP0140",
        "XP0142",
        "XP0146", #температура2
        "XP0147",
        "XP0165",
        "XP0197",
        "XP0200", #параметры
        "XP0201",
        "XP0207",
        "XP0208",
        "XP0209",
        "XP0216",
        "XP0220",
        "XP0234",
        "XP0235",
        "XP0238",
        "XP0239",
        "XP0240",
        "XP0265",
        "XP0297",
        "XP0298",  #параметры
        "XP0301",  #Type - Bridge
        "XP0302",
        "XP0303",  #TDNR
        "XP0304",  #серийный номер весов
        "XP0306",  #Last Service Date
        "XP0308",  #масса калибровочного груза
        "XP0309",
        "XP0312",  #масса калибровочного груза
        "XP0314",  #Cluster ID
        "XP0315",
        "XP0317",
        "XP0320",
        "XP0323",
        "XP0324",
        "XP0325",
        "XP0326",
        "XP0327",
        "XP0328",
        "XP0329",
        "XP0330",
        "XP0331",
        "XP0333",
        "XP0334",
        "XP0335",
        "XP0336",
        "XP0337",
        "XP0338",  #диапазоны
        "XP0339",
        "XP0340",
        "XP0344",
        "XP0345",  #
        "XP0348",
        "XP0354",
        "XP0355",
        "XP0360",
        "XP0361",  #Cell ID
        "XP0362",
        "XP0363",
        "XP0364",
        "XP0366",
        "XP0368",  #масса встроенных грузов
        "XP0369",
        "XP0370",
        "XP0371",
        "XP0372",
        "XP0373",
        "XP0374",
        "XP0375",
        "XP0379",
        "XP0400",  #перезагрузка
        "XP0502",
        "XP0510",
        "XP0600",  #переключение режима
        "XP0601",  #калибровка
        "XP0602",  #внутреняя калибровка
        "XP0603",  #стандартная калибровка
        "XP0605",  #калибровка?
        "XP0606",  #внутреняя калибровка
        "XP0700",
        "XP0701",  #перезагрузка
        "XP1101",
        "XP1107",
        "XP1108",
        "XP1109",
        "XP1120",
        "XP1134",
        "XP1135",
        "XP1138",
        "XP1139",
        "XP1140",
        "XP1165",
        "XP1197",
        "XP1199",
        "XP1700",
        "XP2300",
        "XP2301",
        "XP3100"
    ]

    BaseCommands = [
        "@",
        "I0",
        "I1",
        "I2",
        "I3",
        "I4",
        "I5",
        "S",
        "SI",
        "SIR",
        "Z",
        "ZI",
        "@",
        "D",
        "DW",
        "K",
        "SR",
        "T",
        "TA",
        "TAC",
        "TI",
        "C0",
        "C1",
        "C2",
        "C3",
        "DAT",
        "I10",
        "I11",
        "I14",
        "I15",
        "I16",
        "I17",
        "I18",
        "I19",
        "I20",
        "I21",
        "I22",
        "I23",
        "I24",
        "I25",
        "I26",
        "I27",
        "I28",
        "I29",
        "I51",  #
        "M01",
        "M02",
        "M03",
        "M04",
        "M05",
        "M06",
        "M08",
        "M09",
        "M11",
        "M12",
        "M13",
        "M14",
        "M15",
        "M16",
        "M17",
        "M18",
        "M19",
        "M20",
        "M21",
        "M22",
        "M23",
        "M24",
        "M25",
        "M26",
        "M27",
        "M28",
        "M29",
        "M31",
        "M32",
        "M33",
        "M34",
        "M35",
        "M36",
        "M38",
        "M39",
        "M43",
        "M47",
        "M48",
        "M50",
        "M51",
        "M52",
        "M53",
        "M54",
        "M55",
        "M56",
        "M57",
        "M58",
        "M64",
        "M66",
        "M95",
        "P100",
        "P101",
        "P102",
        "P120",
        "P121",
        "P122",
        "P123",
        "P124",
        "PWR",
        "SIS",
        "SIU",
        "SIUM",
        "SIRU",
        "SNR",
        "SNRU",
        "SRU",
        "ST",
        "SU",
        "SUM",
        "TIM",
        "TST0",
        "TST1",
        "TST2",
        "TST3",
        "UPD",
        "A01",
        "A02",
        "A03",
        "A10",
        "A30",
        "LX20",
        "LX30",
        "LX32",
        "LX39",
        "LX49",
        "LX51",
        "LX52",
        "LX53",
        "LX54",
        "LX112",
        "LX113",
        "LX114",
        "PW"
    ]

    CommandSWaitingAnswer = BaseCommands + XP + XM + XA + XU + ZZ

    def __init__(self, name):
        Thread.__init__(self)
        self.name = name

    def run(self):
        TempNow = 0
        TempPrev = 0
        TimeNow = 0
        TimeLastCommand = datetime.datetime.now().timestamp()
        TimePrev = 0
        try:
            #self.scale.set_buffer_size(rx_size=16384, tx_size=8192)
            self.virtual = serial.Serial(self.virtualName)
            self.scale = serial.Serial(self.serialName)
            self.scale.timeout = 0.5
            self.virtual.timeout = 0.5
            self.Connected = True
            self.scale.write(b'XP0000 4712\r\n')
            time.sleep(1)
            answer = str(self.scale.read_all().decode(encoding="ascii", errors="ignore"))[:-2]
            print(answer)
            if answer != "":
                if(answer.split(' ')[0] == 'ES' or answer.split(' ')[1] != 'A'):
                    self.scale.write(b'XP0000 4712\r\n')
                    time.sleep(1)
                    answer = str(self.scale.read_all().decode(encoding="ascii", errors="ignore"))[:-2]
            while self.Connected:
                virtual2scale = self.virtual.readline()
                virtual2scaleCommand = str(virtual2scale.decode(encoding="ascii", errors="ignore"))[:-2]
                virtualCommandLen = len(virtual2scaleCommand.split(' '))
                virtual2scaleCommand = virtual2scaleCommand.split(' ')[0]
                scale2virtual = self.scale.readline()
                scale2virtualAnswer = str(scale2virtual.decode(encoding="ascii", errors="ignore"))[:-2]
                scale2virtualAnswer = scale2virtualAnswer.split(' ')[0]
                if scale2virtualAnswer in self.AnswersFromScale:
                    self.virtual.write(scale2virtual)
                # if virtual2scale != b'':
                #     print("raw to scale: " + str(virtual2scale))
                #     TimeLastCommand = datetime.datetime.now().timestamp()
                # if datetime.datetime.now().timestamp() - TimeLastCommand > 10:
                #     self.scale.write(b'XP0126\r\n')
                #     TempNow = self.scale.readline().decode(encoding='ascii', errors='ignore')[:-2]
                #     TempNow = str(TempNow).split(' ')
                #     if TempNow[0] == 'XP0126':
                #         TempNow = float(TempNow[3])
                #         TimeNow = datetime.datetime.now().timestamp()
                #
                #         if (TimeNow - TimePrev) != 0 and TimePrev > 0:
                #             dt = (TempNow - TempPrev)/(TimeNow - TimePrev)
                #             print(str(datetime.datetime.now().time()) + ", Т весов = " + str(TempNow) + " \u2103, \u0394T весов, м\u2103/мин: " + str(dt*1000*60))
                #         else:
                #             pass
                #         TempPrev = TempNow
                #         TimePrev = TimeNow
                #         TimeLastCommand = datetime.datetime.now().timestamp()

                # if virtual2scaleCommand in self.Commands1Line:
                #     print("1 to scale: " + str(virtual2scale))
                #     self.scale.write(virtual2scale)
                #     time.sleep(0.1)
                #     scale2virtual = self.scale.readline()
                #     print("1 from scale: " + str(scale2virtual))
                #     self.virtual.write(scale2virtual)

                if virtual2scaleCommand in self.CommandSWaitingAnswer:
                    print("waiting to scale: " + str(virtual2scale))
                    #print(self.scale.get_settings())
                    time_sleep = 0

                    if virtual2scaleCommand == "XP0600":
                        self.scale.timeout = 0.5
                        self.virtual.timeout = 0.5
                        time_sleep = 10
                    elif virtual2scaleCommand == "XP0700":
                        self.scale.timeout = 0.5
                        self.virtual.timeout = 0.5
                        time_sleep = 1
                    elif virtual2scaleCommand == "XU0003":
                        self.scale.timeout = 1
                        self.virtual.timeout = 1
                        time_sleep = 10
                    elif virtual2scaleCommand == "S":
                        self.scale.timeout = 1
                        self.virtual.timeout = 1
                        time_sleep = 1
                    elif virtual2scaleCommand == "SIU":
                        self.scale.timeout = 1
                        self.virtual.timeout = 1
                        time_sleep = 10
                    else:
                        self.scale.timeout = 0.1
                        self.virtual.timeout = 0.1
                        time_sleep = 0.1

                    self.scale.write(virtual2scale)
                    time.sleep(time_sleep)
                    scale2virtual = b''
                    #while self.scale.inWaiting() > 0:
                    scale2virtual = self.scale.readall()
                    self.virtual.write(scale2virtual)
                    print("waiting from scale: " + str(scale2virtual))
                    # while self.scale.inWaiting():
                    #     print("waiting")
                    #     scale2virtual = self.scale.read_all()
                    #     print(scale2virtual)
                    #     time.sleep(0.1)

                # if virtual2scaleCommand in self.CommandsNLine:
                #     self.scale.write(virtual2scale)
                #     time.sleep(0.1)
                #     scale2virtual = b'I0'
                #     self.scale.timeout = 0.1
                #     self.virtual.timeout = 0.1
                #     if virtualCommandLen < 3:
                #         while scale2virtual != b'':
                #             scale2virtual = self.scale.readline()
                #             print("from scale: " + str(scale2virtual))
                #             self.virtual.write(scale2virtual)
                #     else:
                #         scale2virtual = self.scale.read_all()
                #         print("from scale: " + str(scale2virtual))
                #         self.virtual.write(scale2virtual)

                if virtual2scaleCommand in self.CommandTemp:
                    self.scale.write(virtual2scale)
                    time.sleep(1)
                    scale2virtual = self.scale.read_all()
                    print("from scale: " + str(scale2virtual))
                    self.virtual.write(scale2virtual)
                    self.scale.write(b'XP0126\r\n')
                    TempNow = self.scale.readline().decode(encoding='ascii', errors='ignore')[:-2]
                    TempNow = str(TempNow).split(' ')
                    if TempNow[0] == 'XP0126':
                        TempNow = float(TempNow[3])
                        TimeNow = datetime.datetime.now().timestamp()
                        print(TempNow)
                        if (TimeNow - TimePrev) != 0:
                            dt = (TempNow - TempPrev)/(TimeNow - TimePrev)
                            print("\u0394T весов, м\u2103/мин: " + str(dt*1000*60))
                        else:
                            print("0")
                        TempPrev = TempNow
                        TimePrev = TimeNow
        except:
            self.Connected = False
            print("Кажется что-то пошло не так...")


def create_threads():
    name = "SICS Thread"
    sics_tread = SICS(name)
    sics_tread.start()


if __name__ == "__main__":
    create_threads()
