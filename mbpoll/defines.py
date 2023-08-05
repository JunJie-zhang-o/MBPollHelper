from enum import Enum


class Connection():
    SERIAL_PORT = 0
    TCP_IP = 1
    UDP_IP = 2
    ASCII_RTU_OVER_TCP_IP = 3
    ASCII_RTU_OVER_UDP_IP = 4


class BaudRate():
    _300 = 300
    _600 = 600
    _1200 = 1200
    _2400 = 2400
    _4800 = 4800
    _9600 = 9600
    _14400 = 14400
    _19200 = 19200
    _38400 = 38400
    _56000 = 56000
    _115200 = 115200
    _128000 = 128000
    _153600 = 153600
    _230400 = 230400
    _256000 = 256000
    _460800 = 460800
    _921600 = 921600


class DataBits():
    SEVEN = 7
    EIGHT = 8


class Parity():
    NONE = 0
    ODD = 1
    EVEN = 2


class StopBits():
    ONE = 1
    TWO = 2


class Mode():
    RTU = 0
    ASCII = 1


class IPVersion():
    IPV4 = 4
    IPV6 = 6


class DataFormat(Enum):
    # todo:统一格式
    SIGNED = 0
    UNSIGNED = 1
    HEX = 2
    BINARY = 3
    FLOAT_LITTLE_ENDIAN_BYTE_SWAP = 4
    FLOAT_BIG_ENDIAN = 5
    DOUBLE_LITTLE_ENDIAN_BYTE_SWAP = 6
    DOUBLE_BIG_ENDIAN = 7
    SIGNED_32bit_LITTLE_ENDIAN_BYTE_SWAP = 8
    SIGNED_32bit_BIG_ENDIAN = 9

    Floa_little_endian = 10
    Float_big_endian_byte_swap = 11
    Double_little_endian = 12
    Double_big_endian_byte_swap = 13
    Signed_32b_little_endian = 14
    Signed_32bbig_endian_byte_swap = 15
    Unsigned_32b_big_endian = 17
    Unsigned_32b_little_endian_byte_swap = 18
    Unsigned_32b_big_endian_byte_swap = 19
    Unsigned_32b_little_endian = 20
    Signed_64b_big_endian = 21
    Signed_64b_little_endian_byte_swap = 22
    Signed_64b_big_endian_byte_swap = 23
    Signed_64b_little_endian = 24
    Unsigned_64b_big_endian = 25
    Unsigned_64b_little_endian_byte_swap = 26
    Unsigned_64b_big_endian_byte_swap = 27
    Unsigned_64b_little_endian = 28


class ReadWriteResult(Enum):
    # todo:有一些枚举值不对
    SUCCESS = 0
    TIMEOUT_ERROR = 1
    CRC_ERROR = 2
    RESPONSE_ERROR = 3
    WRITE_ERROR = 4
    READ_ERROR = 5
    PORT_NOT_OPEN_ERROR = 16
    DATA_UNINITIALIZED = 10
    INSUFFICIENT_BYTES_RECEIVED = 11
    BYTE_COUNT_ERROR = 16
    TRANSACTION_ID_ERROR = 19
    ILLEGAL_FUNCTION = 81
    ILLEGAL_DATA_ADDRESS = 82
    ILLEGAL_DATA_VALUE = 83
    SERVER_DEVICE_FAILURE = 84
    ACKNOWLEDGE = 85
    SERVER_DEVICE_BUSY = 86
    NAK_NEGATIVE_ACKNOWLEDGMENT = 87
    GATEWAY_PATH_UNAVAILABLE = 810
    GATEWAY_TARGET_DEVICE_FAILED_TO_RESPOND = 811
    

class OpenConnectionResult(Enum):
    SUCCESS = 0
    SERIAL_PORT_NOT_AVAILABLE = 1
    SERIAL_PORT_NOT_POSSIBLE_TO_GET_CURRENT_SETTINGS_FROM_THE_PORT_DRIVE = 3
    SERIAL_PORT__SERIAL_PORT_DRIVE_DID_NOT_ACCEPT_PORT_SETTINGS = 4
    SERIAL_PORT__SERIAL_PORT_DRIVE_DID_NOT_ACCECPT_TIMEOUT_SETTINGS = 5
    TCP_UDP_CONNECTION_FAILED_WSA_START_UP = 12
    TCP_UDP_CONNECTION_FAILED_CONNECT_ERROR = 13
    TCP_UDP_CONNECTION_FAILED_TIMEOUT = 14
    TCP_UDP_CONNECTION_FAILED_IOCTL = 15
    TCP_UDP_CONNECTION_FAILED_SOCKET_ERROR = 17
    TCP_UDP_CONNECTION_FAILED_ADDRESS_INFOMATION = 21
    CONNECTION_ALREADY_OPEN = 255
    

class OpenConnectionError(Exception):
    ...
    

class TimeoutErroe(Exception):
    ...


class ResponseErroe(Exception):
    ...


class CRCError(Exception):
    ...


class WriteError(Exception):
    ...


class ReadError(Exception):
    ...


class InsufficientBytesReceived(Exception):
    ...


class ByteCountError(Exception):
    ...


class TransactionIDError(Exception):
    ...


class IncorrectAddressError(Exception):
    ...


class FunctionCode():
    READ_COILS               = 0x01
    READ_DISCRETE_INPUTS     = 0x02
    READ_HOLDING_REGISTERS   = 0x03
    READ_INPUT_REGISTERS     = 0x04
    WRITE_SINGLE_COIL        = 0x05
    WRITE_SINGLE_REGISTER    = 0x06
    WRITE_MULTIPLE_COILS     = 0x15
    WRITE_MULTIPLE_REGISTERS = 0x16
    