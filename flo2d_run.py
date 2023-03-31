# -*-coding:utf-8 -*-
import win32api
import win32gui
import win32con
import win32ui
import win32process
import time
import sys, os
import shutil
import threading
from multiprocessing import Process


def findWd(label):
    hld = win32gui.FindWindow(None, label)
    if hld > 0:
        return True
    return False


def closeWd(label):
    hld = win32gui.FindWindow(None, label)
    if (hld > 0):
        wnd = win32ui.FindWindow(None, label)
        wnd.SendMessage(win32con.WM_CLOSE)
        return 1
    return 0


def listener():
    while (1):
        time.sleep(50)

        # flag = closeWd("Continue Previous Simulation")
        # flag = closeWd("Warning")

        flag = 0
        flag = closeWd("Simulation Summary")
        flag2 = findWd("VOLUME CONSERVATION")
        if flag or not flag2:
            break
    return


def runmodel(exe, mpath):
    # os.chdir(mpath)
    # os.system('FLOPRO.exe')
    # api参考https://blog.csdn.net/weixin_30437481/article/details/98609470
    win32process.CreateProcess(exe, '', None, None, 0, win32process.CREATE_NO_WINDOW, None, mpath,
                               win32process.STARTUPINFO())
    return


#复制可执行文件函数，移植需修改
def copyExe(destpath):
    srcDir = os.path.abspath('../../exes/FLO-2D/')
    for f in os.listdir(srcDir):
        fDir = srcDir + '\\' + f
        shutil.copy(fDir, destpath)
    return


def clearExe(path):
    for f in os.listdir(path):
        if '.exe' in f or '.dll' in f:
            os.remove(path + '\\' + f)
    return


def main(exe,mpath):
    # exe = os.path.abspath('../exes/FLO-2D/FLOPRO.exe')
    # exe = mpath+'\\FLOPRO.exe'
    # copyExe(mpath)

    # t1 = threading.Thread(target=runmodel,args=(mpath,))
    # t1.start()
    # p1 = Process(target=listener)
    # p1.start()
    # win32api.ShellExecute(0,'open','D:\\Gary\\test\\dist\\close.exe','','',1)
    runmodel(exe, mpath)
    listener()
    # win32process.CreateProcess('D:\\Gary\\test\\dist\\close.exe','',None,None,0,win32process.CREATE_NO_WINDOW,None,None,win32process.STARTUPINFO())

    # clearExe(mpath)
    return


if __name__ == '__main__':
    exePath = os.path.abspath('D:/Example Projects/Flo2D单独运行文件/FLOPRO.exe')
    modelPath = os.path.abspath('D:/Example Projects/Lesson 2 Pro')
    main(exePath, modelPath)
