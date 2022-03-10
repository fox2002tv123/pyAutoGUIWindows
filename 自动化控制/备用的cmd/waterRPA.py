import pyautogui
import time
import xlrd
import pyperclip

# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

# 移动鼠标的函数
def mouseClick(clickTimes, lOrR, img, reTry):
    # 点击次数，[左，右，滚动滑轮]，图片位置，重复次数
    if reTry == 1:
        while True:
            # 循环查找图片位置
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            # 定位的是图片中央位置
            if location is not None:
                # 假如找到图片位置
                pyautogui.click(location.x, location.y, clicks=clickTimes,
                                interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
    # 当时-1时会一直点击（只循环找图片，点击这一步）
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes,
                                interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
    # 当大于1时，点击N次（只循环找图片，点击这一步）
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes,
                                interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查
    i = 1  # i 是行号 [0]是第一列，[1]是第二列
    while i < sheet1.nrows:
        # 第1列 操作类型检查-命令类型检查
        cmdType = sheet1.row(i)[0]
        # 只要excel第一列不是数字，并且不是1-6，就会校验不通过
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('第', i+1, "行,第1列数据有毛病")
            checkCmd = False
            
        # 第2列 内容检查-命令值检查
        cmdValue = sheet1.row(i)[1]
        # 校验cmdType1,2,3鼠标事件
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
            # 图片地址必须为字符串
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空-
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd

# 任务


def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            # 默认运行一次，如果excel第三列,N大于1并类型是INT,运行N次
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            # 移动鼠标到图片位置
            mouseClick(1, "left", img, reTry)
            print("单击左键", img)
        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("右键", img)
        # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", inputValue)
        # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
            
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('欢迎使用不高兴就喝水牌RPA~2.0')
    # 数据检查
    checkCmd = dataCheck(sheet1)
    # 校验通过后
    if checkCmd:
        key = input('选择功能: 1.循环+等待 2.循环到死 3.做一次\n')
        if key == '3':
            # 循环拿出每一行指令
            mainWork(sheet1)
        elif key == '2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
        elif key == '1':
            # 循环+判断等待
            while True:

                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
                key2 = input('选择功能: 1.再做一次 2.退出\n')
                if key2 == '2':
                    break

    else:
        print('输入有误或者已经退出!')
