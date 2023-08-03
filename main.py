

from mbpoll import MBPollTCPApp



app = MBPollTCPApp(ipAddress="172.16.11.31")
# app.closeConnection()
app.openConnection()

doc = app.createReadCoilsDoc(1, 528, 16)
doc.showWindow()
import time
while 1:
    time.sleep(1)
    # values = [doc.getAddress(528+i).value for i in range(16)]
    # values = doc.
    values = doc.readCoils(528, 16)
    print(values)
    print(doc.getRxCount())
    print(doc.getTxCount())
    doc.showWindow()
    app.showCommunicationTraffic()