# -*- coding:utf-8  -*-
# @Time     : 2022/8/4 22:38
# @Author   : BGLB
# @Software : PyCharm
import os
import time

from core import TaoGongChang


def main():
    os.system('taskkill /f /im chromedriver.exe')

    print('***********************操作说明***************************'.replace('*', ' '))
    print('********************第一步: 输入账号***********************'.replace('*', ' '))
    print('********************第二步: 输入密码***********************'.replace('*', ' '))
    print('********************第三步: 手动登录***********************'.replace('*', ' '))
    print('********************第四步: 等待程序自己退出*****************'.replace('*', ' '))
    print('*********************************************************'.replace('*', ' '))
    print('***********************注意事项****************************'.replace('*', ' '))
    print('********************1. 不要手动关闭浏览器*******************'.replace('*', ' '))
    print('********************2. 不要手闭本窗口**********************'.replace('*', ' '))
    print('********************3. 必须安装谷歌浏览器*******************'.replace('*', ' '))
    print('********************下面开始执行*******************'.replace('*', ' '))
    time.sleep(2)
    # import ctypes
    # kernel32 = ctypes.windll.kernel32
    # kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), 128)
    all_task = [TaoGongChang, ]
    for task in all_task:
        task.start()


if __name__ == '__main__':
    main()
