

from mbpoll import MBPollTCPApp



app = MBPollTCPApp(ipAddress="172.16.11.31")
# app.closeConnection()
app.openConnection()

# doc = app.createReadCoilsDoc(1, 528, 16)
doc = app.createMbPollDoc(1)
doc.showWindow()
import time
while 1:
    time.sleep(1)
    doc.writeSingleRegister(33, 5, True)
    # print(values)
    print(doc.getRxCount())
    print(doc.getTxCount())
    print(doc.readResult())
    print(doc.writeResult())
    doc.showWindow()
    app.showCommunicationTraffic()