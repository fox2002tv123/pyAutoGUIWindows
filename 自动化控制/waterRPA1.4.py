import pyautogui
import time
import xlrd
import pyperclip

# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159


def mouseClick(clickTimes, lOrR, img, reTry):
    #
    if reTry == 1:
        i = 1  # 添加循环多少次停止功能
        if '|' not in img:  # 正常情况
            i = 1
            while True:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clickTimes,
                                    interval=0.2, duration=0.2, button=lOrR)
                    break
                print("未找到匹配图片,0.1秒后重试")
                time.sleep(0.1)

                if i > 30:  # 该模块是用于当一直找不到图片，是否跳过当前步骤-附加功能
                    key2 = input('取消当前步骤: 1 是 2 否\n')
                    # if key2 == '1':
                    if key2 == '1' or key2 == '': # 3秒跳出-可是响应''了
                        print(f'跳过步骤 {img}，{lOrR} click')
                        break

                i += 1

        if '|' in img:
            i = 1
            # pass
            arr = img.split(sep='|')  # 由|拆分
            ok_time = 0  # 多图选择次数
            for img in arr:  # 每个都查询下-待改进

                while True:
                    location = pyautogui.locateCenterOnScreen(
                        img, confidence=0.9)
                    if location is not None:
                        pyautogui.click(location.x, location.y, clicks=clickTimes,
                                        interval=0.2, duration=0.2, button=lOrR)
                        ok_time = 1
                        break

                    print("未找到匹配图片,0.1秒后重试")
                    time.sleep(0.1)
                    i += 1
                    if i > 10:  # 该模块是用于当一直找不到图片，是否跳过当前步骤-附加功能
                        # key2 = input('取消当前步骤: 1 是 2 否\n')
                        # if key2 == '1':
                        #     print(f'跳过步骤 {img}，{lOrR} click')
                        #     break
                        break  # 三秒之后跳出查找图片

                if ok_time == 1:  # 当运行了一次跳出多图片循环查找
                    break

                # if i > 30:  # 该模块是用于当一直找不到图片，是否跳过当前步骤-附加功能
                #     key2 = input('取消当前步骤: 1 是 2 否\n')
                #     if key2 == '1':
                #         print(f'跳过步骤 {img}，{lOrR} click')
                #         break

                # i += 1

        # if '|' not in img:  # 正常情况-逻辑错误-可以删除
        #     i = 1
        #     while True:
        #         location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        #         if location is not None:
        #             pyautogui.click(location.x, location.y, clicks=clickTimes,
        #                             interval=0.2, duration=0.2, button=lOrR)
        #             break
        #         print("未找到匹配图片,0.1秒后重试")
        #         time.sleep(0.1)

        #         if i > 30:  # 该模块是用于当一直找不到图片，是否跳过当前步骤-附加功能
        #             key2 = input('取消当前步骤: 1 是 2 否\n')
        #             if key2 == '1':
        #                 print(f'跳过步骤 {img}，{lOrR} click')
        #                 break

        #         i += 1

    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes,
                                interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
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
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮 7.0 判断是否继续
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
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 and cmdType.value != 7.0):
            print('第', i+1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        skipCheck = sheet1.row(i)[3]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i+1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
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
        # 注释(第4列)，内容必须为数字1或者是空-暂时不校验
        # if skipCheck.value != 1.0 or skipCheck.value !=None:
        #     print('第',i+1,"行,第4列数据有毛病")
        #     checkCmd = False
        i += 1
    return checkCmd

# 任务


def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        # 此功能用于第4列有数字1时，跳过该行-注释功能
        skipType = sheet1.row(i)[3]

        if skipType.value == 1.0:
            # 此功能用于第4列有数字1时，跳过该行-注释功能
            while True:
                i += 1
                skipType = sheet1.row(i)[3]
                if skipType.value != 1.0:  # 不是1.0 退出
                    break

        if i < sheet1.nrows:
            #  避免i溢出
            cmdType = sheet1.row(i)[0]
            # 重新加载避免位置错开
            if cmdType.value == 1.0:
                # 取图片名称
                img = sheet1.row(i)[1].value
                reTry = 1
                if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                    reTry = sheet1.row(i)[2].value
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
                # pyautogui.click(x=None,y=None,clicks=1,interval=0.1)
                time.sleep(0.5)
                pyautogui.click(x=None, y=None, clicks=1, interval=0.1)  # 内置单击
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                print("输入:", inputValue)

            # #4代表输入
            # elif cmdType.value == 4.0:
            #     inputValue = sheet1.row(i)[1].value
            #     # pyperclip.copy(inputValue)
            #     # pyautogui.hotkey('ctrl','v')
            #     # time.sleep(0.5)
            #     pyautogui.click(x=None,y=None,clicks=1,interval=0.1)
            #     time.sleep(0.5)
            #     pyautogui.write(inputValue,interval=0.25)
            #     print("输入:",inputValue)
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
            # 7代表判断是否继续-添加功能
            elif cmdType.value == 7.0:
                #
                key_end = input('直接退出: 1.否 2.是\n')
                if key_end == '2':
                    i = sheet1.nrows  # i直接拉满,跳过循环
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # file = r'C:\Users\Administrator\Desktop\waterApp\cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    names = wb.sheet_names()

    sheet_index = int(input(f'选择需要的工作表{list(enumerate( names))}\n:'))
    print(names[sheet_index])
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_name(sheet_name=names[sheet_index])
    # sheet1 = wb.sheet_by_index(0)
    print('欢迎使用不高兴就喝水牌RPA~')
    # 数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key = input('选择功能: 1.循环+选择 2.调试模式 3.运行一次 4.无限循环\n')
        if key == '3':
            # 循环拿出每一行指令
            mainWork(sheet1)
        elif key == '4':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")

        # elif key == '1':  # 添加功能，回车也能执行
        elif key == '1' or key == '':
            # 正常循环+选择模式
            while True:

                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
                # 判断是否需要退出
                key1 = input('选择功能: 1.继续循环 2.退出\n')
                if key1 == '2':
                    break

        elif key == '2':
            # 调试模式
            while True:
                # 每一次都从新读取一遍工作簿
                wb = xlrd.open_workbook(filename=file)

                # 通过索引获取表格sheet页-再次校验
                sheet1 = wb.sheet_by_name(sheet_name=names[sheet_index])
                # sheet1 = wb.sheet_by_index(0)
                checkCmd = dataCheck(sheet1)

                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
                # 判断是否需要退出
                key1 = input('选择功能: 1.继续 2.退出\n')
                if key1 == '2':
                    break

    else:
        print('输入有误或者已经退出!')
