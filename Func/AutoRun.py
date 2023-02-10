from AutoGui import *
import AutoGui
import json
import keyboard as KB

def write_templog(message):
    with open('./config/temp_log.txt', 'a+')as templog:
        templog.write(message)

def on_press():
    global run_flag
    if run_flag:
        pt = str("暂停")
        write_templog(pt + '\n')
        run_flag = not run_flag
    else:
        pt = str("开始")
        write_templog(pt + '\n')
        run_flag = True


def mouseClick(clickTimes, lOrR, img, reTry, ts_k):
    if reTry == 1:
        while True:
            if run_flag:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    success_flag = True
                    break
                else:
                    pt = str(f"未找到匹配图片{img},等待{ts_k}秒")
                    write_templog(pt + '\n')
                    time.sleep(ts_k)
                    success_flag = False
                    break
            else:
                time.sleep(ts_k)

    elif reTry == -1:
        while True:
            if run_flag:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    success_flag = True
                    break
                else:
                    pt = str(f"未找到匹配图片{img},等待{ts_k}秒")
                    write_templog(pt + '\n')
                    time.sleep(ts_k)
                    success_flag = False
                    break
            else:
                time.sleep(ts_k)

    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            if run_flag:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    pt = str("重复")
                    write_templog(pt + '\n')
                    i += 1
                    success_flag = True
                else:
                    pt = str(f"未找到匹配图片{img},等待{ts_k}秒")
                    write_templog(pt + '\n')
                    time.sleep(ts_k)
                    success_flag = False
                    break
            else:
                time.sleep(ts_k)
    return success_flag

# 定义热键事件

# hotkey_get方法用来判断热键组合个数，并把热键传到对应的变量上newinput[0],[1],[2],[3]……[9]只写了10个后续可以添加。
def hotkey_get(hk_g_inputValue):
    newinput = hk_g_inputValue.split(',')
    if len(newinput) == 1:
        pyautogui.hotkey(hk_g_inputValue)
    elif len(newinput) == 2:
        pyautogui.hotkey(newinput[0], newinput[1])
    elif len(newinput) == 3:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2])
    elif len(newinput) == 4:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3])
    elif len(newinput) == 4:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3])
    elif len(newinput) == 5:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4])
    elif len(newinput) == 6:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4], newinput[5])
    elif len(newinput) == 7:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4], newinput[5], newinput[6])
    elif len(newinput) == 8:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4], newinput[5], newinput[6],
                         newinput[7])
    elif len(newinput) == 9:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4], newinput[5], newinput[6],
                         newinput[7], newinput[8])
    elif len(newinput) == 10:
        pyautogui.hotkey(newinput[0], newinput[1], newinput[2], newinput[3], newinput[4], newinput[5], newinput[6],
                         newinput[7], newinput[8], newinput[9])

    # hotkey_Group方法调用hotkey_get方法，并判断其热键内容是否需要循环。


def hotkeyGroup(hotkey_reTry, hkg_inputValue):
    if hotkey_reTry == 1:
        hotkey_get(hkg_inputValue)
        pt = str("执行了：{}".format(hkg_inputValue))

        write_templog(pt + '\n')
        time.sleep(0.1)
    elif hotkey_reTry == -1:
        while True:
            hotkey_get(hkg_inputValue)
            pt = str("执行了：{}".format(hkg_inputValue))

            write_templog(pt + '\n')
            time.sleep(0.1)
    elif hotkey_reTry > 1:
        i = 1
        while i < hotkey_reTry + 1:
            hotkey_get(hkg_inputValue)
            pt = str("执行了：{}".format(hkg_inputValue))

            write_templog(pt + '\n')
            i += 1
            time.sleep(0.1)


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# 7.0 热键组合（最多4个）
# 8.0 粘贴当前时间
# 9.0 系统命令集
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
        pt = str('没有数据。')

        write_templog(pt + '\n')
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0
                                  and cmdType.value != 7.0 and cmdType.value != 8.0 and cmdType.value != 9.0):
            pt = str('第{}行,错误。'.format(i + 1))

            write_templog(pt + '\n')
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 7.0 热键组合，内容不能为空
        if cmdType.value == 7.0:
            if cmdValue.ctype == 0:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 8.0 时间，内容不能为空
        if cmdType.value == 8.0:
            if cmdValue.ctype == 0:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        # 9.0 系统命令集模式，内容不能为空
        if cmdType.value == 9.0:
            if cmdValue.ctype == 0:
                pt = str('第{}行,错误。'.format(i + 1))

                write_templog(pt + '\n')
                checkCmd = False
        i += 1
    return checkCmd


# 任务
def mainWork(sheet1, sheet2, sheet3, ts_k):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = './picture/' + sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            pt = str("单击左键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(1, "left", img, reTry, ts_k)
            if sheet2 is not None and is_success is False:
                success_flag = auxiliray_mainWork(sheet2, ts_k)
                if sheet3 is not None and success_flag != 0:
                    success_flag = auxiliray_mainWork(sheet3, ts_k)
                i += success_flag
            elif is_success is False and sheet2 is None:
                i += -1

        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = './picture/' + sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            pt = str("双击左键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(2, "left", img, reTry, ts_k)
            if sheet2 is not None and is_success is False:
                success_flag = auxiliray_mainWork(sheet2, ts_k)
                if sheet3 is not None and success_flag != 0:
                    success_flag = auxiliray_mainWork(sheet3, ts_k)
                i += success_flag
            elif is_success is False and sheet2 is None:
                i += -1

        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = './picture/' + sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            pt = str("单击右键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(1, "right", img, reTry, ts_k)
            if sheet2 is not None and is_success is False:
                success_flag = auxiliray_mainWork(sheet2, ts_k)
                if sheet3 is not None and success_flag != 0:
                    success_flag = auxiliray_mainWork(sheet3, ts_k)
                i += success_flag
            elif is_success is False and sheet2 is None:
                i += -1

        # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pt = str("输入：{}".format(inputValue))
            write_templog(pt + '\n')
        # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            pt = str("等待{}秒".format(waitTime))
            write_templog(pt + '\n')
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            pt = str("滚轮滑动{}距离".format(int(scroll)))
            write_templog(pt + '\n')
        # 7代表_热键组合
        elif cmdType.value == 7.0:
            # 取重试次数,并循环。
            hotkey_reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                hotkey_reTry = sheet1.row(i)[2].value
            inputValue = sheet1.row(i)[1].value
            hotkeyGroup(hotkey_reTry, inputValue)
            time.sleep(0.2)
        # 8代表_粘贴当前时间
        elif cmdType.value == 8.0:
            # 设置本机当前时间。
            localtime = time.strftime("%Y-%m-%d %H：%M：%S", time.localtime())
            pyperclip.copy(localtime)
            pyautogui.hotkey('ctrl', 'v')
            pt = str("粘贴了本机时间{}:".format(localtime))

            write_templog(pt + '\n')
            time.sleep(0.2)
        # 9代表_系统命令集模式
        elif cmdType.value == 9.0:
            wincmd = sheet1.row(i)[1].value
            os.system(wincmd)
            pt = str("运行系统命令{}:".format(wincmd))
            write_templog(pt + '\n')
            time.sleep(0.2)
        i += 1

def auxiliray_mainWork(sheet, ts_k):
    i = 1
    while i < sheet.nrows:
        # 取本行指令的操作类型
        cmdType = sheet.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = './picture/' + sheet.row(i)[1].value
            reTry = 1
            if sheet.row(i)[2].ctype == 2 and sheet.row(i)[2].value != 0:
                reTry = sheet.row(i)[2].value
            pt = str("单击左键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(1, "left", img, reTry, ts_k)
            if is_success is False:
                break

        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = './picture/' + sheet.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet.row(i)[2].ctype == 2 and sheet.row(i)[2].value != 0:
                reTry = sheet.row(i)[2].value
            pt = str("双击左键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(2, "left", img, reTry, ts_k)
            if is_success is False:
                break

        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = './picture/' + sheet.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet.row(i)[2].ctype == 2 and sheet.row(i)[2].value != 0:
                reTry = sheet.row(i)[2].value
            pt = str("单击右键：{}".format(img))
            write_templog(pt + '\n')
            is_success = mouseClick(1, "right", img, reTry, ts_k)
            if is_success is False:
                break

        # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pt = str("输入：{}".format(inputValue))
            write_templog(pt + '\n')
        # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet.row(i)[1].value
            time.sleep(waitTime)
            pt = str("等待{}秒".format(waitTime))
            write_templog(pt + '\n')
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet.row(i)[1].value
            pyautogui.scroll(int(scroll))
            pt = str("滚轮滑动{}距离".format(int(scroll)))
            write_templog(pt + '\n')
        # 7代表_热键组合
        elif cmdType.value == 7.0:
            # 取重试次数,并循环。
            hotkey_reTry = 1
            if sheet.row(i)[2].ctype == 2 and sheet.row(i)[2].value != 0:
                hotkey_reTry = sheet.row(i)[2].value
            inputValue = sheet.row(i)[1].value
            hotkeyGroup(hotkey_reTry, inputValue)
            time.sleep(0.2)
        # 8代表_粘贴当前时间
        elif cmdType.value == 8.0:
            # 设置本机当前时间。
            localtime = time.strftime("%Y-%m-%d %H：%M：%S", time.localtime())
            pyperclip.copy(localtime)
            pyautogui.hotkey('ctrl', 'v')
            pt = str("粘贴了本机时间{}:".format(localtime))

            write_templog(pt + '\n')
            time.sleep(0.2)
        # 9代表_系统命令集模式
        elif cmdType.value == 9.0:
            wincmd = sheet.row(i)[1].value
            os.system(wincmd)
            pt = str("运行系统命令{}:".format(wincmd))
            write_templog(pt + '\n')
            time.sleep(0.2)
        i += 1
    if i == sheet.nrows:
        return 0
    else:
        return -1


def use_mainWork(ts_var, x_filename, num, x_filename2=None, x_filename3=None):
    with open('./config/Quick_key.json', 'r')as f:
        quick_key = json.load(f)
    KB.add_hotkey(quick_key, on_press)

    global run_flag
    run_flag = True
    ts_k = float(ts_var)

    file = x_filename
    wb = xlrd.open_workbook(filename=file)
    sheet1 = wb.sheet_by_index(0)
    checkCmd = dataCheck(sheet1)

    if x_filename2 is not None:
        file2 = x_filename2
        wb2 = xlrd.open_workbook(filename=file2)
        sheet2 = wb2.sheet_by_index(0)
        checkCmd = dataCheck(sheet2)
    else:
        sheet2 = None

    if x_filename3 is not None:
        file3 = x_filename3
        wb3 = xlrd.open_workbook(filename=file3)
        sheet3 = wb3.sheet_by_index(0)
        checkCmd = dataCheck(sheet3)
    else:
        sheet3 = None

    if checkCmd:
        item = 0
        while item < num:
            mainWork(sheet1, sheet2, sheet3, ts_k)
            time.sleep(ts_k)
            item += 1
    messagebox.showinfo(message='脚本运行结束。')


def loop_use_mainWork(ts_var, x_filename, num,  x_filename2=None, x_filename3=None):
    with open('./config/Quick_key.json', 'r')as f:
        quick_key = json.load(f)
    KB.add_hotkey(quick_key, on_press)

    global run_flag
    run_flag = True
    ts_k = float(ts_var)

    file = x_filename
    wb = xlrd.open_workbook(filename=file)
    sheet1 = wb.sheet_by_index(0)
    checkCmd = dataCheck(sheet1)

    if x_filename2 is not None:
        file2 = x_filename2
        wb2 = xlrd.open_workbook(filename=file2)
        sheet2 = wb2.sheet_by_index(0)
        checkCmd = dataCheck(sheet2)
    else:
        sheet2 = None

    if x_filename3 is not None:
        file3 = x_filename3
        wb3 = xlrd.open_workbook(filename=file3)
        sheet3 = wb3.sheet_by_index(0)
        checkCmd = dataCheck(sheet3)
    else:
        sheet3 = None

    if checkCmd:
        while True:
            mainWork(sheet1, sheet2, sheet3, ts_k)
            time.sleep(ts_k)


