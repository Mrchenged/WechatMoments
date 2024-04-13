import threading
import time
import traceback
import pyautogui
import pyperclip
import win32con
import win32gui
from retry import retry
from win32api import GetSystemMetrics

import log


class AutoScrollSingle(threading.Thread):

    def __init__(self, gui, search_hwnd, friend_name):
        super().__init__()
        self.gui = gui
        self.search_hwnd = search_hwnd
        self.friend_name = friend_name
        self.scrolling = False
        self.resolutions = ['', '1920', '1600', '2560_125', '2560_175', '2560_100', '1366']

    @retry(tries=5, delay=2)
    def find_element(self, image_name):
        element = None
        for resolution in self.resolutions:
            try:
                element = pyautogui.locateCenterOnScreen(
                    f'resource/auto_gui/{resolution}/{image_name}.png',
                    grayscale=True, confidence=0.8)
                if element:
                    break
            except Exception as e:
                log.LOG.warn(f"Can't find {image_name} in resolution: {resolution}")
                pass

        if element is None:
            raise Exception(f"Can't find {image_name}")
        return element

    def find_moments_tab(self):
        return self.find_element('moments_tab')

    def find_search_button(self):
        return self.find_element('search_button')

    def find_friends(self):
        return self.find_element('friends')

    def find_complete(self):
        return self.find_element('complete')

    def set_foreground_window(self, hwnd):
        win32gui.SetForegroundWindow(hwnd)
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
        win32gui.SetWindowPos(hwnd, None, 100, 100, 0, 0, win32con.SWP_NOSIZE)

    def search_friend(self):
        pyperclip.copy(self.friend_name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)

    def scroll(self):
        while self.scrolling:
            try:
                element = self.find_search_button()
                right_bottom = (element.left + element.width, element.top + element.height + 300)
                pyautogui.scroll(-120)
                pyautogui.click(right_bottom)
                time.sleep(0.2)

                search_hwnd = win32gui.FindWindow('Chrome_WidgetWin_0', '微信')
                moments_hwnd = win32gui.FindWindow('SnsWnd', '朋友圈')

                if search_hwnd and moments_hwnd:
                    width = GetSystemMetrics(0)
                    win32gui.SetWindowPos(moments_hwnd, None, 50, 50, 0, 0, win32con.SWP_NOSIZE)
                    win32gui.SetWindowPos(search_hwnd, None, 50, 50, 0, 0, win32con.SWP_NOSIZE)

            except Exception:
                traceback.print_exc()

    def run(self):
        self.scrolling = True

        try:
            search_hwnd = win32gui.FindWindow('Chrome_WidgetWin_0', '微信')
            wechat_hwnd = win32gui.FindWindow('WeChatMainWndForPC', '微信')

            self.set_foreground_window(wechat_hwnd)
            time.sleep(0.3)
            self.set_foreground_window(search_hwnd)

            x, y = self.find_moments_tab()
            pyautogui.click(x, y)
            time.sleep(0.1)

            element = self.find_search_button()
            pyautogui.click(element.left - 100, element.top + element.height / 2)
            time.sleep(0.25)

            pyautogui.write('1')
            time.sleep(0.25)
            pyautogui.click(element.left + element.width / 2, element.top + element.height / 2)
            time.sleep(1.5)

            x, y = self.find_friends()
            pyautogui.click(x, y)
            time.sleep(0.5)

            self.search_friend()

            x, y = self.find_complete()
            pyautogui.click(x, y)
            time.sleep(0.25)

            element = self.find_search_button()
            pyautogui.click(element.left - 100, element.top + element.height / 2)
            time.sleep(0.25)

            pyautogui.press('backspace')
            time.sleep(0.1)
            pyautogui.press('backspace')
            time.sleep(0.25)
            pyperclip.copy('？')
            time.sleep(0.25)

            element = self.find_search_button()
            pyautogui.click(element.left + element.width / 2, element.top + element.height / 2)
            time.sleep(1.0)

            self.scroll()

        except Exception:
            traceback.print_exc()
