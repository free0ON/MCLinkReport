import serial
import time
import datetime
from threading import Thread

class SICS(Thread):
    virtualName = 'COM7'
    serialName = 'COM1'
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
         'XP0000', 'XP0314', 'XP0126', 'XP0700', 'XP0366', 'XP0336', 'XP0306', 'XP0361', 'XP0324', 'XP1700', 'XP2300', 'XP0323', 'XP0344', 'XP0345',      'XP0600',
         'XA0000', 'XA0001', 'XA1005',
         'ZZ00', 'ZZ01',
         'XU0000', 'XU0002', 'XU0003', 'XU1008',  'COPT', 'XD11', 'WS']
    CommandsNLine = ['I0']
    CommandTemp = ['I4', 'I14']
    AnswersFromScale = ['S', ]
    CommandSWaitingAnswer = ['S', 'Z', 'SU']

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
            virtual = serial.Serial(self.virtualName)
            scale = serial.Serial(self.serialName)
            scale.timeout = 0.1
            virtual.timeout = 0.1
            self.Connected = True
            self.scale.write(b'XP0000 4712\r\n')
            time.sleep(1)
            self.scale.read_all()
            while self.Connected:
                virtual2scale = virtual.readline()
                virtual2scaleCommand = str(virtual2scale.decode(encoding="ascii", errors="ignore"))[:-2]
                virtualCommandLen = len(virtual2scaleCommand.split(' '))
                virtual2scaleCommand = virtual2scaleCommand.split(' ')[0]
                scale2virtual = scale.readline()
                scale2virtualAnswer = str(scale2virtual.decode(encoding="ascii", errors="ignore"))[:-2]
                scale2virtualAnswer = scale2virtualAnswer.split(' ')[0]
                if scale2virtualAnswer in self.AnswersFromScale:
                    virtual.write(scale2virtual)
                if virtual2scale != b'':
                    print("raw to scale: " + str(virtual2scale))
                    TimeLastCommand = datetime.datetime.now().timestamp()
                if datetime.datetime.now().timestamp() - TimeLastCommand > 10:
                    scale.write(b'XP0126\r\n')
                    TempNow = scale.readline().decode(encoding='ascii', errors='ignore')[:-2]
                    TempNow = str(TempNow).split(' ')
                    if TempNow[0] == 'XP0126':
                        TempNow = float(TempNow[3])
                        TimeNow = datetime.datetime.now().timestamp()

                        if (TimeNow - TimePrev) != 0 and TimePrev > 0:
                            dt = (TempNow - TempPrev)/(TimeNow - TimePrev)
                            print(str(datetime.datetime.now().time()) + ", Т весов = " + str(TempNow) + " \u2103" +
                                  ", \u0394T весов, м\u2103/мин: " + str(dt*1000*60))
                        else:
                            pass
                        TempPrev = TempNow
                        TimePrev = TimeNow
                        TimeLastCommand = datetime.datetime.now().timestamp()

                if virtual2scaleCommand in self.Commands1Line:
                    print("to scale: " + str(virtual2scale))
                    scale.write(virtual2scale)
                    #time.sleep(0.1)
                    scale2virtual = scale.read_all()
                    print("from scale: " + str(scale2virtual))
                    virtual.write(scale2virtual)

                if virtual2scaleCommand in self.CommandSWaitingAnswer:
                    print("to scale: " + str(virtual2scale))
                    scale.write(virtual2scale)
                    time.sleep(0.1)
                    while scale.inWaiting():
                        scale2virtual = scale.read_all()
                        print(scale2virtual)
                        time.sleep(0.1)
                    print("from scale: " + str(scale2virtual))
                    virtual.write(scale2virtual)


                if virtual2scaleCommand in self.CommandsNLine:
                    scale.write(virtual2scale)
                    time.sleep(0.1)
                    scale2virtual = b'I0'
                    if virtualCommandLen < 2:
                        while scale2virtual != b'':
                            scale2virtual = scale.readline()
                            print("from scale: " + str(scale2virtual))
                            virtual.write(scale2virtual)
                    else:
                        scale2virtual = scale.read_all()
                        print("from scale: " + str(scale2virtual))
                        virtual.write(scale2virtual)

                if virtual2scaleCommand in self.CommandTemp:
                    scale.write(virtual2scale)
                    time.sleep(1)
                    scale2virtual = scale.read_all()
                    print("from scale: " + str(scale2virtual))
                    virtual.write(scale2virtual)
                    scale.write(b'XP0126\r\n')
                    TempNow = scale.readline().decode(encoding='ascii', errors='ignore')[:-2]
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
