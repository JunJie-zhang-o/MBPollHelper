

import sys, os

from mbpoll.defines import DataFormat
sys.path.append(os.getcwd())
from mbpoll import MBPollTCPApp

class TestMBPDoc:
    
    
    def setup_method(self):
        print()
        self.app = MBPollTCPApp(ipAddress="172.16.11.31")
        self.app.openConnection()
        print("++++++++++++++++++++++++++++++++++++++++")

    def teardown_method(self):
        print()
        self.app.closeConnection()
        print("++++++++++++++++++++++++++++++++++++++++")


    def test_write_coils(self):
        doc = self.app.createMbPollDoc(1,100)
        doc.writeMultipleCoils(528, [0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,0])
        doc.writeSingleCoil(559,1)



    def test_read_coils(self):
        doc = self.app.createMbPollDoc(1,100)
        ret = doc.readCoils(528, 32)
        print(ret)
        assert ret == [0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1]


    def test_read_discrete_inputs(self):
        doc = self.app.createMbPollDoc(1,100)
        ret = doc.readDiscreteInputs(528, 16)
        print(ret)
        assert ret == [0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1]
        

    def test_write_multiple_registers(self):
        doc = self.app.createMbPollDoc(1, 100)
        doc.writeMultipleRegisters(33, [0x1234, 0xDB98, 0x3149, 0xAC7A])
    
    def test_read_multiple_registers(self):
        doc = self.app.createMbPollDoc(1, 100)
        ret = doc.readHoldingRegisters(33, 4, format=DataFormat.SIGNED)
        print(ret)
        assert ret == [4660, -9320, 12617, -21382]
        ret = doc.readHoldingRegisters(33, 4, format=DataFormat.UNSIGNED)
        print(ret)
        assert ret == [0x1234, 0xDB98, 0x3149, 0xAC7A]
        ret = doc.readHoldingRegisters(33, 4, format=DataFormat.FLOAT_BIG_ENDIAN)
        print(ret)
        assert ret == [5.706865537029703e-28, 2.934739118387597e-09]
