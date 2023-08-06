

from mbpoll import MBPollTCPApp
from mbpoll import DataFormat



app = MBPollTCPApp(ipAddress="172.16.11.31")
app.openConnection()

doc = app.createMbPollDoc(1,10)
doc1 = app.createMbPollDoc(1,10)
doc2 = app.createMbPollDoc(1,10)
doc3 = app.createMbPollDoc(1,10)

doc1.readHoldingRegisters(31, 5)
doc1.showWindow()
doc1.setRows()
doc1.resizeWindow()
doc1.readHoldingRegisters(31, 5)
import time
while 1:
    # time.sleep(1)
    # doc.writeSingleRegister(33, 5, True)
    # # print(values)
    # print(doc.getRxCount())
    # print(doc.getTxCount())
    # print(doc.readResult())
    # print(doc.writeResult())
    # doc.showWindow()
    # app.showCommunicationTraffic()
    a = doc1.readHoldingRegisters(31, 5, format=DataFormat.UNSIGNED)
    print(a)
    doc1.writeMultipleRegisters(33,[0x12345678],format=DataFormat.UNSIGNED_32B_LITTLE_ENDIAN, isCreateWin=False)
    