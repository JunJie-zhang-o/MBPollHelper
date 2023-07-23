

from mbpoll import MBPollTCPApp



app = MBPollTCPApp(ipAddress="172.16.11.31")
# app.closeConnection()
app.openConnection()



import time
while 1:
    time.sleep(1)