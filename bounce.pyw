# import pywintypes
import win32gui
import win32con
import win32com
import win32com.client
import win32api
import pythoncom
import threading
import time
import random
import win32ui

username = win32api.GetUserName()


def msgbox():
    win32api.Beep(1050, 1000)
    desktophwnd = win32gui.GetDesktopWindow()
    win32gui.MessageBox(desktophwnd, "Hello, " + username + "!",
                        "Congratulations!", win32con.MB_ICONERROR)


def greet():
    pythoncom.CoInitialize()
    s = win32com.client.Dispatch('SAPI.SpVoice')
    voices = [i for i in s.GetVoices()]
    s.Voice = voices[1]
    s.Rate = 1
    s.Speak(f"Congratulations, {username}")


class Wndw:
    def __init__(self, hwnd, starting_speed, rect, bounds):
        self.hwnd = hwnd
        self.starting_rect = rect
        self.left_x, self.top_y, self.right_x, self.bottom_y = rect
        self.width = self.right_x - self.left_x
        self.height = self.bottom_y - self.top_y
        self.speed_x, self.speed_y = starting_speed
        self.bounds = bounds

    @property
    def rect(self):
        return (self.left_x, self.top_y, self.right_x, self.bottom_y)

    def move_me(self, gravity):
        bounds = self.bounds
        self.speed_y += gravity
        if (self.right_x + self.speed_x >= bounds[2] and self.speed_x > 0) or (self.left_x - self.speed_x < bounds[0] and self.speed_x < 0):
            self.speed_x *= -1
        else:
            self.left_x += self.speed_x
            self.right_x += self.speed_x

        if (self.bottom_y + self.speed_y >= bounds[3] and self.speed_y > 0) or (self.top_y - self.speed_y < bounds[1] and self.speed_y < 0):
            self.speed_y *= -1
        else:
            self.top_y += self.speed_y
            self.bottom_y += self.speed_y

        if not self.reset_if_outside():
            self.update_position()

    def update_position(self):
        win32gui.MoveWindow(self.hwnd, int(self.left_x), int(
            self.top_y), self.width, self.height, 1)

    def reset_if_outside(self):
        if outside(self.rect, self.bounds):
            self.reset_to_start()
            return True
        return False

    def reset_to_start(self):
        self.left_x, self.top_y, self.right_x, self.bottom_y = self.starting_rect
        self.update_position()


def outside(rect1, rect2):
    return rect1[2] < rect2[0] or rect1[0] > rect2[2] or rect1[1] > rect2[3] or rect1[3] < rect2[1]


def window_enumeration_handler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


def bounce_around(window, gravity, fps):
    while True:
        try:
            window.move_me(gravity)
            time.sleep(1 / fps)
        except Exception:
            break


def get_randomized_speed(min_speed, max_speed):
    speed_abs = min_speed + random.random() * (max_speed - min_speed)
    return speed_abs * random.choice([1, -1])


def init_window(hwnd, min_speed, max_speed):
    monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromWindow(hwnd))
    speed_x = get_randomized_speed(min_speed, max_speed)
    speed_y = get_randomized_speed(min_speed, max_speed)
    window = Wndw(hwnd, (speed_x, speed_y),
                  win32gui.GetWindowRect(hwnd), monitor_info["Work"])
    return window


threading.Thread(target=msgbox).start()
threading.Thread(target=greet).start()

time.sleep(2)

msgboxhwnd = win32ui.FindWindow(None, "Congratulations!").GetSafeHwnd()
msgbox = init_window(msgboxhwnd, 10, 10)
bounce_around(msgbox, 4, 60)
win32com.client.Dispatch('SAPI.SpVoice').Speak("See you again soon.")
