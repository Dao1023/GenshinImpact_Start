import cv2 
import pyautogui
import os
import time
import subprocess
import win32com.client
import numpy as np
from PIL import ImageGrab

while True:
    # 检查原神是否已经启动
    if os.system('tasklist /FI "IMAGENAME eq YuanShen.exe" 2>NUL | find /I /N "YuanShen.exe">NUL') == 0:
        print("原神 已在运行!")
        break
        
    # 获取屏幕分辨率
    screen_width, screen_height = pyautogui.size()  
    
    # 截图
    print("正在检测屏幕...")
    screenshot = cv2.cvtColor(np.array(ImageGrab.grab(bbox=(0,0,screen_width,screen_height))), cv2.COLOR_BGR2RGB)

    # 计算屏幕白色像素比例
    white_pixels = np.count_nonzero(screenshot == [255, 255, 255])
    total_pixels = screenshot.shape[0] * screenshot.shape[1]
    white_percentage = white_pixels / total_pixels * 100

    # 判断是否满足启动条件
    if white_percentage >= 90:
        try:
            # 获取快捷方式路径
            shortcut = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\原神\原神.lnk'  
            
            # 解析快捷方式获取安装路径
            shell = win32com.client.Dispatch("WScript.Shell")  
            install_dir = shell.CreateShortCut(shortcut)
            install_dir = install_dir.TargetPath.replace('launcher.exe', '')

            
            # 拼接游戏exe路径 
            game_exe = os.path.join(install_dir, 'Genshin Impact Game', 'YuanShen.exe')
            
            # 将游戏置顶启动
            subprocess.Popen(game_exe)
            subprocess.Popen(['ffplay', '-nodisp', '-autoexit', 'C:\\Users\\Dao\\Projects\\有趣的项目\\GenshinImpact_Start\\mp3\\mp3\\BGM.mp3'])
            
            
            max_retries = 8
            retries = 0

            while retries < max_retries:
                try:
                    time.sleep(1)
                    
                    # 枚举窗口,找到名称包含"原神"的窗口
                    window = pyautogui.getWindowsWithTitle("原神")[0]

                    # 将目标窗口置顶
                    # 容易失败需要执行多次
                    for i in range(1000):
                        pyautogui.moveTo(window.left, window.top)

                    break  # 如果命令成功执行，跳出循环

                except Exception as e:  # 捕获所有异常，或者你可以指定某种异常
                    print(f"Error: {e}")
                    retries += 1
                    if retries >= max_retries:
                        print("Max retries reached!")
                        break
            print("原神 启动!")
            break            
        except:
            print("未获取到原神安装路径!或权限不够!")
            break