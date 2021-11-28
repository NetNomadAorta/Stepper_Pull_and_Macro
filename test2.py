import time
import os
import sys
from subprocess import Popen, PIPE, STDOUT


# User parameters to set: 
SLEEP_TIME_BETWEEN = 1 
SLEEP_TIME_END = 5 



script_path = 'test.py'

p = Popen([sys.executable, '-u', script_path], 
          stdout=PIPE, 
          stdin=PIPE,
          stderr=STDOUT, 
          bufsize=1)

with p.stdout:
    for line in iter(p.stdout.readline, b''):
        line = str(line).replace("b'", "").replace("\\r\\n","").replace("'", "")
        print(line)
        if "place" in line:
            time.sleep(SLEEP_TIME_BETWEEN)
            p.stdin.write("OKAY".encode())
            p.stdin.close()
p.wait()


time.sleep(SLEEP_TIME_END)