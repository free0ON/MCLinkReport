import serial
import time
import datetime
from threading import Thread


class Klimet(Thread):
    KlimetPortName = 'COM7'
    KlimetPort = serial.Serial()
    Connected = False

    def run(self):
        self.KlimetPort = serial.Serial(self.KlimetPortName)
        self.KlimetPort.baudrate = 2400
        #self.KlimetPort.timeout = 2
        #self.KlimetPort.write_timeout = 1
        self.Connected = True
        while self.Connected:

            #answer = str(self.KlimetPort.readline().decode(encoding="ascii", errors="ignore"))
            answer = self.KlimetPort.read_all()
            if answer != b'': print(answer)
            if answer == b'AI\r':
                print(' get AI and answer')
                self.KlimetPort.write(r'KLIMET A30, Ver 1.3, 00, SN 000002\r'.encode(encoding="ascii"))

if __name__ == "__main__":
    klimet_tread = Klimet()
    klimet_tread.start()