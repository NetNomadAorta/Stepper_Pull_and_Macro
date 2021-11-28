import sys, subprocess, time

# User parameters to set: 
SLEEP_TIME_BETWEEN = 2 
SLEEP_TIME_END = 5 

script_path = 'test.py'

class Console():
    def __init__(self):
        if '-r' not in sys.argv:
            self.p = subprocess.Popen(
                [sys.executable, '-u', script_path],
                stdin=subprocess.PIPE,
                # stdout=subprocess.PIPE,
                creationflags=subprocess.CREATE_NEW_CONSOLE
                )
        else:
            while True:
                data = sys.stdin.read(1)
                if not data:
                    break
                sys.stdout.write(data)

    def write(self, data):
        self.p.stdin.write(data.encode('utf8'))
        self.p.stdin.flush()
    
    # def read(self):
    #     test = self.p.stdout
    #     with test:
    #         for line in iter(test.readline, b''):
    #             line = str(line).replace("b'", "").replace("\\r\\n","").replace("'", "")
    #             print(line)
    #     p.wait()

if (__name__ == '__main__'):
    p = Console()
    if '-r' not in sys.argv:
        for i in range(0, 2):
            # p.read()
            p.write("Okay\n")
            time.sleep(SLEEP_TIME_BETWEEN)

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