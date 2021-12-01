import time
from pynput.keyboard import Key, Controller


# User parameters to set: 
SLEEP_TIME_BETWEEN = 0.5 
SLEEP_TIME_END = 5 


def time_convert(sec):
  mins = sec // 60
  sec = sec % 60
  hours = mins // 60
  mins = mins % 60
  print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours) ,int(mins), round(sec)))


# Presses Alt + Tab
def altTab():
    time.sleep(SLEEP_TIME_BETWEEN)
    with keyboard.pressed(Key.alt):
        SLEEP_TIME_BETWEEN
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)


# Type's in string
def typeThis(toType):
    time.sleep(0.25)
    keyboard.type(toType)
    time.sleep(0.25)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)



# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n")

keyboard = Controller()

while True:
    test = input("Press enter to continue.")
    
    altTab()
    
    typeThis("Hello World")


# Starting stopwatch to see how long process takes
print("\n\nTotal Time: ")
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)