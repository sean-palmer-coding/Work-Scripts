import win32api
import time
import math

while True:
    for i in range(500):
        x = int(1000+math.sin(math.pi*i/100)*1000)
        y = int(1000+math.cos(i)*300)
        win32api.SetCursorPos((x, y))
        time.sleep(.01)
    time.sleep(10)
