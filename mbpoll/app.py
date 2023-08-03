import atexit

import win32com.client as win32

from mbpoll.defines import *
from mbpoll.document import MBPollDocument


class MBPollBaseApp:
    """
        MBPoll python 自动化 App基类
    """
    MBPOLL_APP_NAME = "Mbpoll.Application"


    def __init__(self, connection:Connection, responseTimeout:int=1000, delayBetweenPolls:int=20) -> None:
        self._app = win32.Dispatch(self.MBPOLL_APP_NAME)
        self.connection = connection
        self.responseTimeout = responseTimeout
        self.delayBetweenPolls = delayBetweenPolls
        
    
    def _openConnection(self):
        self._app.Connection = self.connection
        self._app.ResponseTimeout = self.responseTimeout
        self._app.DelayBetweenPolls = self.delayBetweenPolls

        self.closeConnection()
        # todo:Exception
        openResult = self._app.OpenConnection
        if openResult != OpenConnectionResult.SUCCESS.value:
            raise OpenConnectionError(OpenConnectionResult(openResult).name)
        atexit.register(self.closeConnection)


    def closeConnection(self):
        self._app.CloseConnection


    def showCommunicationTraffic(self):
        self._app.ShowCommunicationTraffic


    def closeCommunicationTraffic(self):
        self._app.CloseCommunicationTraffic
        
    
    def createMbPollDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument : 
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        return doc
    
    
    def createReadCoilsDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        doc.readCoils(address, quantity)
        return doc
    
    
    def createReadDiscreteInputsDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        doc.readDiscreteInputs(address, quantity)
        return doc
    

    def createReadInputRegistersDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        doc.readInputRegisters(address, quantity)
        return doc
        
    
    def createWriteSingleCoilDoc(self, slaveID:int, address:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity = 1, scanRate= scanRate)
        return doc
    
    
    def createWriteSingleRegisterDoc(self, slaveID:int, address:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity = 1, scanRate= scanRate)
        return doc
    
    
    def createMultipleCoilsDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        return doc
    
    
    def createMultipleRegistersDoc(self, slaveID:int, address:int, quantity:int, scanRate:int= 100) -> MBPollDocument :
        doc = MBPollDocument(slaveID, address, quantity, scanRate= scanRate)
        return doc


    



class MBPollSerialApp(MBPollBaseApp):
    """
        MBPoll python 自动化 App 串口主站类
    """
    def __init__(
        self,
        serialPort: int,
        baudRate: BaudRate = BaudRate._9600,
        dataBits: DataBits = DataBits.EIGHT,
        parity: Parity = Parity.NONE,
        stopBits: StopBits = StopBits.ONE,
        mode: Mode = Mode.RTU,
        removeEcho: bool = False,
        responseTimeout: int = 1000,
        DelayBetweenPolls: int = 20,
    ) -> None:
        self.serialPort = serialPort
        self.baudRate = baudRate
        self.dataBits = dataBits
        self.parity = parity
        self.mode = mode
        self.stopBits = stopBits
        self.removeEcho = removeEcho
        super().__init__(
            Connection.SERIAL_PORT,
            responseTimeout,
            DelayBetweenPolls,
        )

    def openConnection(self):
        self._app.SerialPort = self.serialPort
        self._app.BaudRate = self.baudRate
        self._app.Parity = self.parity
        self._app.DataBits = self.dataBits
        self._app.StopBits = self.stopBits
        self._app.Mode = self.mode
        self._app.RemoveEcho = self.removeEcho
        self._openConnection()


class MBPollTCPApp(MBPollBaseApp):
    """
        MBPoll python 自动化 App TCP主站类
    """
    def __init__(
        self,
        ipAddress: str = "127.0.0.1",
        port: int = 502,
        connectTimeout: int = 1000,
        responseTimeout: int = 1000,
        DelayBetweenPolls: int = 20,
        ipVersion: IPVersion = IPVersion.IPV4,
        mode: Mode = None,
    ) -> None:
        self.ipAddress = ipAddress
        self.port = port
        self.connectionTimeout = connectTimeout
        self.mode = mode
        self.ipVersion = ipVersion
        if mode is None:
            super().__init__(Connection.TCP_IP, responseTimeout, DelayBetweenPolls)
        else:
            super().__init__(Connection.ASCII_RTU_OVER_TCP_IP, responseTimeout, DelayBetweenPolls)

    def openConnection(self):
        self._app.IPAddress = self.ipAddress
        self._app.ServerPort = self.port
        self._app.IPVersion = self.ipVersion
        self._app.ConnectTimeout = self.connectionTimeout
        if self.mode is not None:
            self._app.Mode = self.mode
        self._openConnection()


class MBPollUDPApp(MBPollBaseApp):
    """
        MBPoll python 自动化 App UDP主站类
    """
    def __init__(
        self,
        ipAddress: str = "127.0.0.1",
        port: int = 502,
        connectTimeout: int = 1000,
        responseTimeout: int = 1000,
        DelayBetweenPolls: int = 20,
        ipVertion: IPVersion = IPVersion.IPV4,
        mode: Mode = None,
    ) -> None:
        self.ipAddress = ipAddress
        self.port = port
        self.connectionTimeout = connectTimeout
        self.mode = mode
        self.ipVertion = ipVertion
        if mode is None:
            super().__init__(Connection.UDP_IP, responseTimeout, DelayBetweenPolls)
        else:
            super().__init__(Connection.ASCII_RTU_OVER_UDP_IP, responseTimeout, DelayBetweenPolls)

    def openConnection(self):
        self._app.IPAddress = self.ipAddress
        self._app.ServerPort = self.port
        self._app.IPVertion = self.ipVertion
        self._app.ConnectTimeout = self.connectionTimeout
        if self.mode is not None:
            self._app.Mode = self.mode
        self._openConnection()
