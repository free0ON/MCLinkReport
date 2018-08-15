import socket
import datetime
con = socket.socket()


con.connect(("192.168.0.1", 8001))
con.send(b"user\n")
con.send(b"admin\n")
con.send(b"r af0151\n")
Feed = con.recv(1024).split(' ')
con.send(b"r af0152\n")
CorrFeed = con.recv(1024).split(' ')
con.send(b"r af0153\n")
CorrTotal = con.recv(1024).split(' ')


con.send(b"contout\n")
data = b""

data = con.recv(1024)
startTime = datetime.datetime.now()
while data:
    data = con.recv(1024)
    out = data.decode("utf-8").split(' ')
    try:
        V = str(out[0]).split("~")[3]
        G = out[1]
        M = out[2]
        T = out[3]
    except:
        V = 0
        G = 0
        M = 0
        T = 0
        pass
    curTime = startTime - datetime.datetime.now()
    print(str(curTime)+";" + str(G) + ";" + str(V) + ";" + str(M) + ";" + str(T))

con.send(b"xgroup all\n")