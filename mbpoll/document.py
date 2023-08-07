
import time
from typing import Union
import win32com.client as win32
from mbpoll.defines import ByteOrder, IncorrectAddressError, DataFormat, FunctionCode, ReadWriteResult

 
class MBPollBaseDocument():
    MBPOLL_DOC_NAME = "Mbpoll.Document"
    
    DEFAULT_SCANRATE = 500  # unit:ms

    def __init__(self, slaveID:int, scanRate:int=DEFAULT_SCANRATE) -> None:
        """init MBPoll Document object

        Args:
            slaveID (int): Modbus Slave ID
            scanRate (int, optional): scanRate. Defaults to DEFAULT_SCANRATE.
        """
        self.__slaveID  = slaveID
        self.__scanRate = scanRate
        self.startingAddress = 0
        self.quantity = 0
        self.__doc = win32.Dispatch(self.MBPOLL_DOC_NAME)
        
    
    @property
    def slaveID(self):
        return self.__slaveID


    @slaveID.setter
    def slaveID(self, id):
        self.__slaveID = id
        
        
    @property
    def scanRate(self):
        return self.__scanRate


    @scanRate.setter
    def scanRate(self, rate):
        self.__scanRate = rate
        
        
    @property
    def document(self):
        return self.__doc
    

    def showWindow(self) -> None:
        """As default Modbus document windows are hidden. 
           The ShowWindow function makes Modbus Poll visible and shows the document with data content.
        """
        self.__doc.ShowWindow()
        
        
    def setRows(self, numberRows=4, autoResizeWindow:bool=False) -> None:
        """Specify the number of rows in the grid

        Args:
            numberRows (int, optional): 0:10 Rows 1:20 Rows 2:50 Rows 3:100 Rows 4:Fit to quantity. Defaults to 4.
            autoResizeWindow (bool, optional): Resize all columns to fit the values inside the cells.. Defaults to False.
        """
        self.__doc.Rows(numberRows)
        if autoResizeWindow:
            self.resizeAllColumns()
    
    
    def resizeWindow(self) -> None:
        """Resize an opened window to fit the grid.
        """
        self.__doc.ResizeWindow()
    
    
    def getTxCount(self) -> int:
        """Gets the number of requests.

        Returns:
            int: The number of requests.
        """
        return self.__doc.GetTxCount
    
    
    def getRxCount(self) -> int:
        """Gets the number of response.

        Returns:
            int: The number of response.
        """
        return self.__doc.GetRxCount
    

    def resizeAllColumns(self) -> None:
        """Resize all columns to fit the values inside the cells.
        """
        self.__doc.ResizeAllColumns()
    
    
    def readResult(self) -> ReadWriteResult:
        """Use this function to check if communication established with Read is running successful.

        Returns:
            ReadWriteResult: ReadResult
        """
        return ReadWriteResult(self.__doc.ReadResult)
    
    
    def writeResult(self) -> ReadWriteResult:
        """Use this function to check if a write was successful.

        Returns:
            ReadWriteResult: WriteResult
        """
        return ReadWriteResult(self.__doc.WriteResult)
    
    
    def enableRefresh(self, enable:bool = True) -> None:
        self.__doc.EnableRefresh = enable

    
    def addressIsValid(self, address) -> None:
        if address>= self.startingAddress and address <= self.startingAddress + self.quantity:
            return True
        else:
            raise IncorrectAddressError()
    
    
    def getName(self, address:int) -> str:
        """Gets the name of a address.

        Args:
            address (int): Modbus address.

        Returns:
            str: Modbus address name
        """
        if self.addressIsValid(address):
            return self.__doc.GetName(address - self.startingAddress)
        
    
    def setName(self, address:int, name:str) -> None:
        """Changes the name of a address

        Args:
            address (int): Modbus address.
            name (str): Modbus address name
        """
        if self.addressIsValid(address):
            self.__doc.SetName(address - self.startingAddress, name)
    

    def formatAll(self, dataFormat:Union[DataFormat, int]) -> None:
        """Format all value cells with the selected format.

        Args:
            dataFormat (Union[DataFormat, int]): The format of the value cell.
        """
        self.__doc.FormatAll(dataFormat)
    
    
    def getFormat(self, address:int) -> DataFormat:
        """Gets the display format of the Modbus value.

        Args:
            address (int): Modbus address

        Returns:
            DataFormat: Modbus address dispaly format
        """
        if self.addressIsValid(address):
            return DataFormat(self.__doc.getFormat(address- self.startingAddress))
    
    
    def setFormat(self, address:int, format:Union[DataFormat, int]) -> None:
        """Change the display format of the Modbus values. See Format values above.

        Args:
            address (int): Modbus address
            format (Union[DataFormat, int]): The format of the value cell.
        """
        if self.addressIsValid(address):
            format = format if type(format) == int else format.value
            self.__doc.SetFormat(address - self.startingAddress, format)


    # def save(self, savePath:str) -> bool:
    #     return self.__doc.Save(savePath)


class MBPollDocument(MBPollBaseDocument):
    MBPOLL_DOC_NAME = "Mbpoll.Document"
    
    def __init__(self, slaveID, scanRate) -> None:
        super().__init__(slaveID, scanRate)
        self.functionCode = None
        # self.__readAddress = None
        
    
    def __getValue(self, address:int, quantity:int, format:DataFormat) -> list:
        if format in [DataFormat.FLOAT_BIG_ENDIAN, DataFormat.DOUBLE_BIG_ENDIAN, DataFormat.SIGNED_32B_BIG_ENDIAN, DataFormat.UNSIGNED_32B_BIG_ENDIAN, DataFormat.SIGNED_64B_BIG_ENDIAN, DataFormat.UNSIGNED_64B_BIG_ENDIAN]:
            self.document.ByteOrder = ByteOrder.BIG_ENDIAN
        elif format in [DataFormat.FLOAT_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP]:
            self.document.ByteOrder = ByteOrder.LITTLE_ENDIAN_BYTE_SWAP
        elif format in [DataFormat.FLOAT_BIG_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_BIG_ENDIAN_BYTE_SWAP]:
            self.document.ByteOrder = ByteOrder.BIG_ENDIAN_BYTE_SWAP
        elif format in [DataFormat.FLOAT_LITTLE_ENDIAN, DataFormat.DOUBLE_LITTLE_ENDIAN, DataFormat.SIGNED_32B_LITTLE_ENDIAN, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN, DataFormat.SIGNED_64B_LITTLE_ENDIAN, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN]:
            self.document.ByteOrder = ByteOrder.LITTLE_ENDIAN
        if format in [DataFormat.SIGNED, DataFormat.HEX, DataFormat.BINARY]:
            values = [self.document.SRegisters(index) for index in range(quantity)]
        elif format == DataFormat.UNSIGNED:
            values = [self.document.URegisters(index) for index in range(quantity)]
        elif format in [DataFormat.FLOAT_BIG_ENDIAN, DataFormat.FLOAT_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_BIG_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_LITTLE_ENDIAN]:
            values = [self.document.Floats(index) for index in range(0, quantity, 2)]
        elif format in [DataFormat.DOUBLE_BIG_ENDIAN, DataFormat.DOUBLE_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_BIG_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_LITTLE_ENDIAN]:
            values = [self.document.Doubles(index) for index in range(0, quantity, 4)]
        elif format in [DataFormat.SIGNED_32B_BIG_ENDIAN, DataFormat.SIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_LITTLE_ENDIAN]:
            values = [self.document.Ints_32(index) for index in range(0, quantity, 2)]
        elif format in [DataFormat.UNSIGNED_32B_BIG_ENDIAN, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN]:
            values = [self.document.UInts_32(index) for index in range(0, quantity, 2)]
        elif format in [DataFormat.SIGNED_64B_BIG_ENDIAN, DataFormat.SIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_LITTLE_ENDIAN]:
            values = [self.document.Ints_64(index) for index in range(0, quantity, 4)]
        elif format in [DataFormat.UNSIGNED_64B_BIG_ENDIAN, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN]:
            values = [self.document.UInts_64(index) for index in range(0, quantity, 4)]
        return values
    
    
    def __setValue(self, address:int, values:list, format:DataFormat, isWinVersion:bool=False) -> None:
        #! ByteOrder 不控制Win地址
        if not isWinVersion:
            if format in [DataFormat.FLOAT_BIG_ENDIAN, DataFormat.DOUBLE_BIG_ENDIAN, DataFormat.SIGNED_32B_BIG_ENDIAN, DataFormat.UNSIGNED_32B_BIG_ENDIAN, DataFormat.SIGNED_64B_BIG_ENDIAN, DataFormat.UNSIGNED_64B_BIG_ENDIAN]:
                self.document.ByteOrder = ByteOrder.BIG_ENDIAN
            elif format in [DataFormat.FLOAT_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP]:
                self.document.ByteOrder = ByteOrder.LITTLE_ENDIAN_BYTE_SWAP
            elif format in [DataFormat.FLOAT_BIG_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_BIG_ENDIAN_BYTE_SWAP]:
                self.document.ByteOrder = ByteOrder.BIG_ENDIAN_BYTE_SWAP
            elif format in [DataFormat.FLOAT_LITTLE_ENDIAN, DataFormat.DOUBLE_LITTLE_ENDIAN, DataFormat.SIGNED_32B_LITTLE_ENDIAN, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN, DataFormat.SIGNED_64B_LITTLE_ENDIAN, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN]:
                self.document.ByteOrder = ByteOrder.LITTLE_ENDIAN

        startIndex = address - self.startingAddress
        startIndex = 0
        if format == DataFormat.SIGNED:
            if isWinVersion:
                [self.document.SRegistersWin(index+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.SRegisters(index+startIndex, values[index]) for index in range(len(values))]
                return len(values)
        elif format in [DataFormat.UNSIGNED, DataFormat.HEX, DataFormat.BINARY]:
            if isWinVersion:
                [self.document.URegistersWin(index+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.URegisters(index+startIndex, values[index]) for index in range(len(values))]
                return len(values)
        elif format in [DataFormat.FLOAT_BIG_ENDIAN, DataFormat.FLOAT_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_BIG_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.FloatsWin(index*2+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.Floats(index*2+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 2
        elif format in [DataFormat.DOUBLE_BIG_ENDIAN, DataFormat.DOUBLE_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_BIG_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.DoublesWin(index*4+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.Doubles(index*4+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 4
        elif format in [DataFormat.SIGNED_32B_BIG_ENDIAN, DataFormat.SIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.Ints_32Win(index*2+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.Ints_32(index*2+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 2
        elif format in [DataFormat.UNSIGNED_32B_BIG_ENDIAN, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.UInts_32Win(index*2+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.UInts_32(index*2+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 2
        elif format in [DataFormat.SIGNED_64B_BIG_ENDIAN, DataFormat.SIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.Ints_64Win(index*4+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.Ints_64(index*4+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 4
        elif format in [DataFormat.UNSIGNED_64B_BIG_ENDIAN, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN]:
            if isWinVersion:
                [self.document.UInts_64Win(index*4+startIndex, values[index]) for index in range(len(values))]
            else:
                [self.document.UInts_64(index*4+startIndex, values[index]) for index in range(len(values))]
                return len(values) * 4
            
    def __getNums(self, values:list, format:DataFormat):
        if format == DataFormat.SIGNED:
            return len(values)
        elif format in [DataFormat.UNSIGNED, DataFormat.HEX, DataFormat.BINARY]:
            return len(values)
        elif format in [DataFormat.FLOAT_BIG_ENDIAN, DataFormat.FLOAT_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_BIG_ENDIAN_BYTE_SWAP, DataFormat.FLOAT_LITTLE_ENDIAN]:
            return len(values) * 2
        elif format in [DataFormat.DOUBLE_BIG_ENDIAN, DataFormat.DOUBLE_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_BIG_ENDIAN_BYTE_SWAP, DataFormat.DOUBLE_LITTLE_ENDIAN]:
            return len(values) * 4
        elif format in [DataFormat.SIGNED_32B_BIG_ENDIAN, DataFormat.SIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_32B_LITTLE_ENDIAN]:
            return len(values) * 2
        elif format in [DataFormat.UNSIGNED_32B_BIG_ENDIAN, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_32B_LITTLE_ENDIAN]:
            return len(values) * 2
        elif format in [DataFormat.SIGNED_64B_BIG_ENDIAN, DataFormat.SIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.SIGNED_64B_LITTLE_ENDIAN]:
            return len(values) * 4
        elif format in [DataFormat.UNSIGNED_64B_BIG_ENDIAN, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_BIG_ENDIAN_BYTE_SWAP, DataFormat.UNSIGNED_64B_LITTLE_ENDIAN]:
            return len(values) * 4
    
    
    def readCoils(self, address:int, quantity:int = 1) -> list:
        """Modbus function code 01

        Args:
            address (int): Modbus address
            quantity (int, optional): The number of data. Defaults to 1.

        Returns:
            list: Data list of corresponding addresses
        """
        if self.functionCode != FunctionCode.READ_COILS or self.startingAddress != address:
            self.document.ReadCoils(self.slaveID, address, quantity, self.scanRate)
            self.functionCode = FunctionCode.READ_COILS
            self.startingAddress = address
            self.quantity = quantity
            time.sleep(self.scanRate/1000)
        result = [self.document.Coils(index) for index in range(quantity)]
        return result
    
    
    def readDiscreteInputs(self, address:int, quantity:int = 1) -> list:
        """Modbus function code 02

        Args:
            address (int): Modbus address
            quantity (int, optional): The number of data. Defaults to 1.

        Returns:
            list: Data list of corresponding addresses
        """
        if self.functionCode != FunctionCode.READ_DISCRETE_INPUTS or self.startingAddress != address:
            self.document.ReadDiscreteInputs(self.slaveID, address, quantity, self.scanRate)
            self.functionCode = FunctionCode.READ_DISCRETE_INPUTS
            self.startingAddress = address
            self.quantity = quantity
            time.sleep(self.scanRate/1000)
        result = [self.document.Coils(index) for index in range(quantity)]
        return result
    
    
    def readHoldingRegisters(self, address:int, quantity:int = 1, format:DataFormat=DataFormat.SIGNED):
        """Modbus function code 03

        Args:
            address (int): Modbus address
            quantity (int, optional): The number of data. Defaults to 1.

        Returns:
            list: Data list of corresponding addresses
        """
        if self.functionCode != FunctionCode.READ_HOLDING_REGISTERS or self.startingAddress != address:
            self.document.ReadHoldingRegisters(self.slaveID, address, quantity, self.scanRate)
            self.functionCode = FunctionCode.READ_HOLDING_REGISTERS
            self.startingAddress = address
            self.quantity = quantity
            time.sleep(self.scanRate/1000)
        result = self.__getValue(address, quantity, format)
        return result
    
    
    def readInputRegisters(self, address:int, quantity:int = 1, format:DataFormat=DataFormat.SIGNED):
        """Modbus function code 04

        Args:
            address (int): Modbus address
            quantity (int, optional): The number of data. Defaults to 1.

        Returns:
            list: Data list of corresponding addresses
        """
        if self.functionCode != FunctionCode.READ_INPUT_REGISTERS or self.startingAddress != address:
            self.document.ReadInputRegisters(self.slaveID, address, quantity, self.scanRate)
            self.functionCode = FunctionCode.READ_INPUT_REGISTERS
            self.startingAddress = address
            self.quantity = quantity
            time.sleep(self.scanRate/1000)
        result = self.__getValue(address, quantity, format)
        return result
    
    
    def writeSingleCoil(self, address:int, data:int, isCreateWin:bool=False, scanRateForWin:int=None, enableRefresh:bool=True):
        """Modbus function code 05

        Args:
            address (int): Modbus address
            data (int): Data  of corresponding addresses
            isCreateWin (bool, optional): Whether to create a write window. Defaults to False.
            scanRateForWin (int, optional): Refresh frequency of the write window, just enable for isCreateWin == True. Defaults to None.
            enableRefresh (bool, optional): Not Use. Defaults to True.
        """
        if isCreateWin:
            scanRate = self.scanRate if scanRateForWin is None else scanRateForWin
            self.document.WriteSingleCoilWin(self.slaveID, address, scanRate)
            self.functionCode = FunctionCode.WRITE_SINGLE_COIL
            self.startingAddress = address
            self.quantity = 1
            self.document.CoilsWin(0, int(bool(data)))
        else:
            self.document.Coils(0 + address - self.startingAddress, int(bool(data)))
            self.document.WriteSingleCoil(self.slaveID, address)
    
    # # todo:写单寄存器时 需要看当前是什么格式，并且只能写入一个寄存器
    # def writeSingleRegister(self, address, data, format:DataFormat=DataFormat.SIGNED, isCreateWin:bool=False, scanRateForWin:int=None, enableRefresh:bool=True):
    #     if isCreateWin:
    #         scanRate = self.scanRate if scanRateForWin is None else scanRateForWin
    #         self.document.WriteSingleRegisterWin(self.slaveID, address, scanRate)
    #         self.functionCode = FunctionCode.READ_INPUT_REGISTERS
    #         self.startingAddress = address
    #         self.quantity = 1
    #         self.__setValue(address, [data], format, isWinVersion=True)
    #     else:
    #         # todo
    #         self.__setValue(address, [data], format, isWinVersion=False)
    #         self.document.WriteSingleRegister(self.slaveID, address)
    
    
    def writeMultipleCoils(self, address:int, data:list, isCreateWin:bool=False, scanRateForWin:int=None, enableRefresh:bool=True):
        """Modbus function code 15

        Args:
            address (int): Modbus address
            data (int): Data list of corresponding addresses
            isCreateWin (bool, optional): Whether to create a write window. Defaults to False.
            scanRateForWin (int, optional): Refresh frequency of the write window, just enable for isCreateWin == True. Defaults to None.
            enableRefresh (bool, optional): Not Use. Defaults to True.
        """
        if isCreateWin:
            scanRate = self.scanRate if scanRateForWin is None else scanRateForWin
            self.document.WriteMultipleCoilsWin(self.slaveID, address, len(data), scanRate)
            self.functionCode = FunctionCode.WRITE_MULTIPLE_COILS
            self.startingAddress = address
            self.quantity = 1
            for index in range(len(data)):
                self.document.CoilsWin(index, int(bool(data[index])))
        else:
            for index in range(len(data)):
                self.document.Coils(index, int(bool(data[index])))
            self.document.WriteMultipleCoils(self.slaveID, address, len(data))
    
    
    def writeMultipleRegisters(self, address:int, data:list, format:DataFormat=DataFormat.SIGNED,isCreateWin:bool=False, scanRateForWin:int=None, enableRefresh:bool=True):
        """Modbus function code 16

        Args:
            address (int): Modbus address
            data (list): Data list of corresponding addresses
            format (DataFormat, optional): _description_. Defaults to DataFormat.SIGNED.
            isCreateWin (bool, optional): Whether to create a write window. Defaults to False.
            scanRateForWin (int, optional): Refresh frequency of the write window, just enable for isCreateWin == True. Defaults to None.
            enableRefresh (bool, optional): Not Use. Defaults to True.
        """
        if isCreateWin:
            scanRate = self.scanRate if scanRateForWin is None else scanRateForWin
            self.document.WriteMultipleRegistersWin(self.slaveID, address, self.__getNums(data,format), scanRate)
            self.functionCode = FunctionCode.WRITE_MULTIPLE_REGISTERS
            self.startingAddress = address
            self.quantity = self.__getNums(data,format)
            self.__setValue(address, data, format, isWinVersion=True)
        else:
            self.__setValue(address, data, format, isWinVersion=False)
            self.document.WriteMultipleRegisters(self.slaveID, address, self.__getNums(data,format))

    
    