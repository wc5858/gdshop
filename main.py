import win32gui
import win32con
import win32api
from PIL import Image, ImageGrab
import time
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files (x86)/Tesseract-OCR/tesseract'

def get_window_info():
    name = 'Grim Dawn'
    handle = win32gui.FindWindow(0, name)  # 获取窗口句柄
    if handle == 0:
        return None
    else:
        window_size = win32gui.GetWindowRect(handle)
        win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, window_size[0], window_size[1], window_size[2] - window_size[0],
                              window_size[3] - window_size[1], win32con.SWP_NOMOVE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW)
        win32api.SetCursorPos((window_size[0] + 30, window_size[1] + 30))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN | win32con.MOUSEEVENTF_LEFTUP,
                             window_size[0] + 30, window_size[1] + 30, 0, 0)
        time.sleep(1)
        return window_size

def click(window_size, x, y):
    win32api.SetCursorPos((window_size[0] + x, window_size[1] + y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, window_size[0] + x, window_size[1] + y, 0, 0)
    time.sleep(0.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, window_size[0] + x, window_size[1] + y, 0, 0)

def r_click(window_size, x, y):
    win32api.SetCursorPos((window_size[0] + x, window_size[1] + y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, window_size[0] + x, window_size[1] + y, 0, 0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, window_size[0] + x, window_size[1] + y, 0, 0)

def world_trans(window_size):
    win32api.keybd_event(77, 0x32, 0, 0) #按m
    win32api.keybd_event(77, 0x32, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
    click(window_size, 720, 786)
    time.sleep(0.2)
    click(window_size, 773, 462)
    time.sleep(20)
    win32api.keybd_event(77, 0x32, 0, 0)  # 按m
    win32api.keybd_event(77, 0x32, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    click(window_size, 720, 786)
    time.sleep(0.1)
    click(window_size, 669, 458)
    time.sleep(10)

def open_store(window_size):
    click(window_size, 740, 38)
    time.sleep(3)
    click(window_size, 743, 381)
    time.sleep(0.1)

def check_buy(res):
    data = [
        ['Overlord', 'Colossal Fortress', 'Thorns'],
        ['Overlord', 'Colossal Fortress', 'Vengeance'],
        ['Sandstorm', 'Fleshwarped Tome', 'Oracle'],
        ['Sandstorm', 'Fleshwarped Tome', 'Rituals'],
        ['Sandstorm', 'Fleshwarped Tome', 'Scorched Runes'],
        ['Sandstorm', 'Fleshwarped Tome', 'the Sands'],
        ['Sandstorm', 'Fleshwarped Tome', 'Wildfire'],
        ['Destroyer', 'Fleshwarped Tome', 'Oracle'],
        ['Destroyer', 'Fleshwarped Tome', 'Rituals'],
        ['Destroyer', 'Fleshwarped Tome', 'the Sands'],
        ['Destroyer', 'Fleshwarped Tome', 'Wildfire'],
        ['Thunderstruck', 'Fleshwarped Tome', 'Oracle'],
        ['Thunderstruck', 'Fleshwarped Tome', 'the Sands'],
        ['Thunderstruck', 'Fleshwarped Tome', 'Torrents'],
        ['Thunderstruck', 'Fleshwarped Tome', 'Scorched Runes']
    ]
    for i in range(len(data)):
        flag = 1
        for j in range(len(data[i])):
            if not(data[i][j] in res):
                flag = 0
        if flag == 1:
            return 1
    return 0

def buy_green(window_size):
    click(window_size, 388, 224)
    time.sleep(0.1)
    store = [[0 for i in range(10)] for j in range(10)]
    for i in range(10):
        last_empty = 0
        for j in range(10):
            if store[i][j] == 1:
                continue
            x = 228 + 17 + j * 32
            y = 221 + 24 + 17 + i * 32
            win32api.SetCursorPos((window_size[0] + x, window_size[1] + y))
            time.sleep(0.3)
            img = ImageGrab.grab((window_size[0], window_size[1]+70, window_size[0] + 1000, window_size[1] + 175))
            # addr = 'temp.png'
            # img.save(addr, 'png')
            res = pytesseract.image_to_string(img, lang='eng')
            if res == '':
                if last_empty == 2:
                    return
                else:
                    last_empty += 1
            else:
                last_empty = 0
            if 'Rare Shield' in res:
                last_empty = 0
                for m in range(3):
                    for n in range(2):
                        if (i + m) < 10 and (j + n) < 10:
                            store[i + m][j + n] = 1
            elif 'Off-Hand' in res:
                last_empty = 0
                for m in range(2):
                    for n in range(2):
                        if (i + m) < 10 and (j + n) < 10:
                            store[i + m][j + n] = 1
            else:
                if last_empty == 2:
                    return
                else:
                    last_empty += 1
            if check_buy(res) == 1:
                r_click(window_size, x, y)
                time.sleep(0.1)

def buy_material(window_size):
    click(window_size, 445, 224)
    time.sleep(0.1)
    store = [[0 for i in range(10)] for j in range(10)]
    for i in range(10):
        last_empty = 0
        for j in range(10):
            if store[i][j] == 1:
                continue
            x = 228 + 17 + j * 32
            y = 221 + 24 + 17 + i * 32
            win32api.SetCursorPos((window_size[0] + x, window_size[1] + y))
            time.sleep(0.3)
            img = ImageGrab.grab((window_size[0] + x + 25, window_size[1]+y-20, window_size[0] + x + 300 , window_size[1] + y + 60))
            # addr = 'temp' + str(i) + str(j) + '.png'
            # img.save(addr, 'png')
            res = pytesseract.image_to_string(img, lang='eng')
            if res == '':
                if last_empty == 1:
                    return
                else:
                    last_empty = 1
            else:
                last_empty = 0
            # print(res)
            if 'enged Plating' in res:
                if not('enged Plating (1)' in res):
                    r_click(window_size, x, y)
            if 'Aether Crystal' in res:
                r_click(window_size, x, y)
            if 'Scrap' in res:
                for m in range(2):
                    for n in range(2):
                        if (i + m) < 10 and (j + n) < 10:
                            store[i + m][j + n] = 1
                r_click(window_size, x, y)
            # if ('Aether Cluster' in res) or ('Aether Shard' in res):
            if ('Aether Shard' in res):
                if i< 9:
                    store[i + 1][j] = 1
                r_click(window_size, x, y)
            time.sleep(0.1)

if __name__=="__main__":
    window_size = get_window_info()
    while(1):
        try:
            world_trans(window_size)
            open_store(window_size)
            buy_green(window_size)
            buy_material(window_size)
        except:
            win32api.keybd_event(27, 1, 0, 0)  # 按m
            time.sleep(0.1)
            win32api.keybd_event(27, 1, win32con.KEYEVENTF_KEYUP, 0)