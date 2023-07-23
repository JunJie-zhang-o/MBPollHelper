
import win32com.client as win32
from mbpoll.defines import IncorrectAddressError, DataFormat





class MBPollAddress():
    
    def __init__(self, docObj, address, index, name:str="", dataFormat:DataFormat = DataFormat.SIGNED) -> None:
        self.__docObj = docObj
        self.__address = address
        self.__index = index
        self.__name = name
        self.__dataFormat = dataFormat
        
        
    @property
    def name(self):
        return self.__name
    
    
    @name.setter
    def name(self, name:str):
        self.__name = name
        self.__docObj.SetName(self.__index, name)
        
        
    @property
    def dataFormat(self):
        return self.__dataFormat
    
    
    @dataFormat.setter
    def dataFormat(self, format:DataFormat):
        self.__dataFormat = format
        self.__docObj.SetFormat(self.__index, format)
        
        
    @property
    def value(self):
        return self.__docObj.SRegisters(self.__index)


    @value.setter
    def value(self, value):
        self.__docObj.SRegisters(self.__index, value)
 


class MBPollDocument():
    MBPOLL_DOC_NAME = "Mbpoll.Document"
    
    def __init__(self, slaveID, address, quantity, scanRate) -> None:
        self.__slaveID  = slaveID
        self.scanRate = scanRate
        self.startingAddress = address
        self.__doc = win32.Dispatch(self.MBPOLL_DOC_NAME)
        self.registers = [MBPollAddress(self.__doc, address=address, index=i) for i in range(quantity)]
        
    
    @property
    def slaveID(self):
        return self.__slaveID


    @slaveID.setter
    def slaveID(self, id):
        self.__slaveID = id
    
    
    def addressIsValid(self, address):
        if address>= self.startingAddress and address <= self.startingAddress + len(self.registers):
            return True
        else:
            raise IncorrectAddressError()

    
    def getAddress(self, address) -> MBPollAddress:
        if self.addressIsValid(address):
            register:MBPollAddress = self.registers[address - self.startingAddress]
            return register

    
    def readCoils(self, address, quantity:int = 1) -> list:
        if self.addressIsValid(address):
            self.__doc.ReadCoils(self.slaveID, address, quantity, self.scanRate)
            result = [self.getAddress(address+i).value for i in range(quantity)]
            return result
    
    
    def readDiscreteInputs(self, address, quantity:int = 1):
        if self.addressIsValid(address):
            self.__doc.ReadDiscreteInputs(self.slaveID, address, quantity, self.scanRate)
            result = [self.getAddress(address+i).value for i in range(quantity)]
            return result
    
    
    def readHoldingRegisters(self, address, quantity:int = 1):
        if self.addressIsValid(address):
            self.__doc.ReadHoldingRegisters(self.slaveID, address, quantity, self.scanRate)
            result = [self.getAddress(address+i).value for i in range(quantity)]
            return result
    
    
    def readInputRegisters(self, address, quantity:int = 1):
        if self.addressIsValid(address):
            self.__doc.ReadInputRegisters(self.slaveID, address, quantity, self.scanRate)
            result = [self.getAddress(address+i).value for i in range(quantity)]
            return result
    
    
    def writeSingleCoil(self, address, data):
        if self.addressIsValid(address):
            coil =  self.getAddress(address)
            coil.value = data
            self.__doc.WriteSingleCoil(self.slaveID, address)
    
    
    def writeSingleRegister(self, address, data):
        if self.addressIsValid(address):
            register =  self.getAddress(address)
            register.value = data
            self.__doc.WriteSingleRegister(self.slaveID, address)
    
    
    def writeMultipleCoils(self, address, data:list):
        if self.addressIsValid(address):
            for index in range(len(data)):
                coil = self.getAddress(address+index)
                coil.value = data[index]
            self.__doc.WriteMultipleCoils(self.slaveID, address)
    
    
    def writeMultipleRegisters(self, address, data):
        if self.addressIsValid(address):
            for index in range(len(data)):
                register = self.getAddress(address+index)
                register.value = data[index]
            self.__doc.WriteMultipleRegisters(self.slaveID, address)


    def writeSingleCoilWin(self, address, data):
        if self.addressIsValid(address):
            coil = self.getAddress(address)
            coil.value = data
            self.__doc.WriteSingleCoilWin(self.slaveID, address, self.scanRate)
    
    
    def writeSingleRegisterWin(self, address, data):
        if self.addressIsValid(address):
            register = self.getAddress(address)
            register.value = data
            self.__doc.WriteSingleRegisterWin(self.slaveID, address, self.scanRate)
    
    
    def writeMultipleCoilsWin(self, address, data):
        if self.addressIsValid(address):
            for index in range(len(data)):
                coil = self.getAddress(address+index)
                coil.value = data
            self.__doc.WriteMultiCoilsWin(self.slaveID, address, self.scanRate)
    
    
    def writeMultipleRegistersWin(self, address, data):
        if self.addressIsValid(address):
            for index in range(len(data)):
                register = self.getAddress(address+index)
                register.value = data
            self.__doc.WriteMultiCoilsWin(self.slaveID, address, self.scanRate)

    
    
    def showWindow(self):
        self.__doc.ShowWindow()
        
        
    def setRows(self, numberRows, autoResizeWindow:bool=False):
        """_summary_

        Args:
            rows (_type_): _description_
            autoResizeWindow (_type_): _description_
        """
        self.__doc.Rows(numberRows)
        if autoResizeWindow:
            self.resizeAllColumns()
    
    
    def resizeWindow(self):
        self.__doc.ResizeWindow()
    
    
    def getTxCount(self):
        self.__doc.GetTxCount()
    
    
    def getRxCount(self):
        self.__doc.GetRxCount()
    
    
    def getAddressName(self, address):
        return self.getAddress(address).name
        
    
    def setAddressName(self, address, name):
        self.getAddress(address).name = name
    

    def formatAll(self, dataFormat:DataFormat):
        self.__doc.FormatAll(dataFormat)
    
    
    def getFormat(self, address):
        return self.getAddress(address).dataFormat
    
    
    def setFormat(self, address, format:DataFormat):
        self.getAddress(address).dataFormat = format
    
    
    def resizeAllColumns(self):
        self.__doc.ResizeAllColumns()
    
    
    def readResult(self):
        return self.__doc.ReadResult()
    
    
    def writeResult(self):
        return self.__doc.WriteREsult()
    
    
    def enableRefresh(self, enable:bool = True):
        self.__doc.EnableRefresh = enable