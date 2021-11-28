import time
import os
import sys
from subprocess import Popen, PIPE, STDOUT


# User parameters to set: 
SLEEP_TIME_BETWEEN = 1 
SLEEP_TIME_END = 3 



script_path = 'test.py'

p = Popen([sys.executable, script_path], 
          stdout=PIPE, 
          stdin=PIPE,
          )

while True:
    stdout = p.stdout.readline()
    print(stdout)
    
    if "place" in str(stdout):
        p.stdin.write("OKAY".encode())
        p.stdin.flush()
        p.stdout.close()
        # print(p.stdout.readline())

# with p.stdout:
#     for line in iter(p.stdout.readline, b''):
#         line = str(line).replace("b'", "").replace("\\r\\n","").replace("'", "")
#         print(line)
#         if "place" in line:
#             time.sleep(SLEEP_TIME_BETWEEN)
#             p.stdin.write("OKAY".encode())
#             p.stdin.close()
# p.wait()


time.sleep(SLEEP_TIME_END)