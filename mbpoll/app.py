import atexit

import win32com.client as win32

from mbpoll.defines import *


class MBPollBaseApp:
    """
        MBPoll python 自动化 App基类
    """
    MBPOLL_APP_NAME = "Mbpoll.Application"

    def __init__(self, connection:Connection, responseTimeout:int=1000, delayBetweenPolls:int=20) -> None:
        self.__app = win32.Dispatch(self.MBPOLL_APP_NAME)
        self.connection = connection
        self.responseTimeout = responseTimeout
        self.delayBetweenPolls = delayBetweenPolls

    def __openConnection(self):
        self.__app.ResponseTimeout = self.responseTimeout
        self.__app.DelayBetweenPolls = self.delayBetweenPolls

        # todo:Exception
        self.__app.OpenConnection
        atexit.register(self.closeConnection)

    def closeConnection(self):
        self.__app.CloseConnection

    def showCommunicationTraffic(self):
        self.__app.ShowCommunicationTraffic

    def closeCommunicationTraffic(self):
        self.__app.CloseCommunicationTraffic


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
        self.__app.SerialPort = self.serialPort
        self.__app.BaudRate = self.baudRate
        self.__app.Parity = self.parity
        self.__app.DataBits = self.dataBits
        self.__app.StopBits = self.stopBits
        self.__app.Mode = self.mode
        self.__app.RemoveEcho = self.removeEcho
        self.__openConnection()


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
        ipVertion: IPVersion = IPVersion.IPV4,
        mode: Mode = None,
    ) -> None:
        self.ipAddress = ipAddress
        self.port = port
        self.connectionTimeout = connectTimeout
        self.mode = mode
        self.ipVertion = ipVertion
        if mode is None:
            super().__init__(Connection.ASCII_RTU_OVER_TCP_IP, responseTimeout, DelayBetweenPolls)
        else:
            super().__init__(Connection.TCP_IP, responseTimeout, DelayBetweenPolls)

    def openConnection(self):
        self.__app.IPAddress = self.ipAddress
        self.__app.ServerPort = self.port
        self.__app.IPVertion = self.ipVertion
        self.__app.ConnectTimeout = self.connectionTimeout
        if self.mode is not None:
            self.__app.Mode = self.mode
        self.__openConnection()


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
            super().__init__(Connection.ASCII_RTU_OVER_UDP_IP, responseTimeout, DelayBetweenPolls)
        else:
            super().__init__(Connection.UDP_IP, responseTimeout, DelayBetweenPolls)

    def openConnection(self):
        self.__app.IPAddress = self.ipAddress
        self.__app.ServerPort = self.port
        self.__app.IPVertion = self.ipVertion
        self.__app.ConnectTimeout = self.connectionTimeout
        if self.mode is not None:
            self.__app.Mode = self.mode
        self.__openConnection()
