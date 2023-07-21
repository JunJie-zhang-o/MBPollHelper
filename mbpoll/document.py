



class MBPollDocument():
    
    
    def __init__(self, slaveID, scanRate) -> None:
        pass
    
    
    def readCoils(self, address, quantity):
        pass
    
    
    def readDiscreteInputs(self, address, quantity):
        pass
    
    def readHoldingRegisters(self, address, quantity):
        pass
    
    def readInputRegisters(self, address, quantity):
        pass
    
    
    def writeSingleCoil(self, address, data):
        pass
    
    
    def writeSingleRegister(self, address, data):
        pass
    
    def writeMultipleCoils(self, address, data):
        pass
    
    
    def writeMultipleRegisters(self, address, data):
        pass
    
    def showWindow(self):
        pass
        
    def setRows(self, rows, autoResizeWindow):
        pass
    
    def resizeWindow(self):
        pass
    
    def getTxCount(self):
        pass
    
    def getRxCount(self):
        pass
    
    def getAddressName(self, index):
        pass
    
    def setAddressName(self, index, name):
        pass
    
    def formatAll(self, dataFormat):
        pass
    
    
    def getFormat(self):
        pass
    
    def setFormat(self):
        pass
    
    def resizeAllColumns(self):
        pass
    
    def readResult(self):
        pass
    
    def writeResult(self):
        pass
    
    
    def EnableRefresh(self):
        pass