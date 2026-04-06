import win32api
import time
import keyboard
run = True
import mouse
from mouse import MoveEvent, ButtonEvent
import cv2


def main():

    while run is True:
        a = win32api.GetKeyState(0x01)  # Left mouse button down
        b = win32api.GetKeyState(0x57)  # [W] Key Pressed
        if a < 0:
            x, y = win32api.GetCursorPos()
            print(x, y)
            time.sleep(0.2)
        if b < 0:
            mouse.play([MoveEvent(x=x, y=y, time=0.2)])
            break
if __name__ == "__main__":
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))