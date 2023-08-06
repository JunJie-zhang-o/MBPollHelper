

import sys, os
sys.path.append(os.getcwd())
import sys
print(sys.path)
from mbpoll import MBPollTCPApp


# class TestMBPDoc:
    
    
# def __init__(self) -> None:
app = MBPollTCPApp(ipAddress="172.16.11.31")
app.openConnection()


def test_write_coils():
    doc = app.createMbPollDoc(1,100)
    doc.writeMultipleCoils(528, [0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1])



def test_read_coils():
    doc = app.createMbPollDoc(1,100)
    ret = doc.readCoils(528, 32)
    assert ret == [0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1,0,1]


