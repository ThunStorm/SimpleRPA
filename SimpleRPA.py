import pyautogui
import time
import xlrd
import pyperclip

# 置于True可自行选择循环或单次执行。False为单次执行
EXEC_NUM_SWITCH = False

class SimpleRPA:

    def __init__(self, script, sheet_num, img_dir):
        self.script = script
        self.img_dir = img_dir
        self.sheet_num = sheet_num
        self.sheet = None
        self.dc_flag = None

    # 数据检查
    # cmd.value  1.0 左键单击    2.0 左键双击    3.0 右键单击    4.0 输入    5.0 等待    6.0 滚轮    7.0 回车    8.0 粘贴
    # ctype     空：0
    #           字符串：1
    #           数字：2
    #           日期：3
    #           布尔：4
    #           error：5
    def data_check(self):
        is_checked = True
        # 行数检查
        if self.sheet.nrows < 2:
            print("未发现可执行的RPA操作")
            is_checked = False

        # 每行数据检查
        for i in range(1, self.sheet.nrows):
            cmd, args = self.sheet.row(i)[0], self.sheet.row(i)[1]
            if cmd.ctype != 2 or (cmd.value not in [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0]):
                print("第{}行第1列，数据出错，请检查".format(i + 1))
                is_checked = False
            elif cmd.value in [1.0, 2.0, 3.0] and args.ctype != 1:
                print("第{}行第2列，参数非字符串类型，请检查".format(i + 1))
                is_checked = False
            elif cmd.value == 4.0 and args.ctype == 0:
                print("第{}行第2列，参数为空，请检查".format(i + 1))
                is_checked = False
            elif cmd.value in [5.0, 6.0] and args.ctype != 2:
                print("第{}行第2列，参数非数字，请检查".format(i + 1))
                is_checked = False
            elif cmd.value == [7.0, 8.0] and args.ctype != 0:
                print("第{}行第2列，参数非空，请检查".format(i + 1))
        return is_checked

    # 定义鼠标事件
    def mouse_click(self, clicks, button, img, redo):
        if redo == 1:
            while True:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clicks, interval=0.2, duration=0.2, button=button)
                    break
                print("未找到匹配图片,0.1秒后重试")
                time.sleep(0.1)
        elif redo == -1:
            while True:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clicks, interval=0.2, duration=0.2, button=button)
                time.sleep(0.1)
        elif redo > 1:
            i = 1
            while i < redo + 1:
                location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                if location is not None:
                    pyautogui.click(location.x, location.y, clicks=clicks, interval=0.2, duration=0.2, button=button)
                    print("操作成功，重复中")
                    i += 1
                time.sleep(0.1)

    # 操作
    def operations(self):
        i = 1
        while i < self.sheet.nrows:
            # 取本行指令的操作类型
            cmd, args, redo_flag = self.sheet.row(i)[0], self.sheet.row(i)[1], self.sheet.row(i)[2]
            # 鼠标操作
            if cmd.value in [1.0, 2.0, 3.0]:
                img = self.img_dir + args.value
                redo = redo_flag.value if (redo_flag.ctype == 2 and redo_flag.value != 0) else 1
                self.mouse_click(1 if int(cmd.value % 2) == 1 else 2, "left" if cmd.value <= 2.0 else "right", img, redo)
                print("对{}区域执行{}操作".format(img, cmd.value))
            # 输入
            elif cmd.value == 4.0:
                inputs = args.value
                pyperclip.copy(inputs)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                print("已输入: {}".format(inputs))
            # 等待
            elif cmd.value == 5.0:
                wait_time = args.value
                time.sleep(wait_time)
                print("等待 {} 秒".format(wait_time))
            # 滚轮
            elif cmd.value == 6.0:
                scroll_dist = args.value
                pyautogui.scroll(int(scroll_dist))
                print("滚轮滑动距离为 {}".format(int(scroll_dist)))
            # 单击回车
            elif cmd.value == 7.0:
                pyautogui.press('enter')
                print("单击ENTER")
            # 粘贴
            elif cmd.value == 8.0:
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                print("已粘贴")

            i += 1

    # 运行程序
    def run(self):
        # 打开文件
        wb = xlrd.open_workbook(filename=self.script)
        self.sheet = wb.sheet_by_index(self.sheet_num)
        self.dc_flag = self.data_check()
        if self.dc_flag:
            #
            key = '1'
            if not EXEC_NUM_SWITCH:
                key = input('选择功能: 1.做一次 2.循环到死 \n')
            if key == '1':
                # 循环拿出每一行指令
                self.operations()
            elif key == '2':
                while True:
                    self.operations()
                    time.sleep(0.1)
                    print("等待0.1秒")
        else:
            print('输入有误或者已经退出!')


if __name__ == '__main__':
    script_dir = './cmd/cmd.xls'
    img_dir = './cmd/img/'
    sheet_NO = 0
    app = SimpleRPA(script_dir, sheet_NO, img_dir)
    app.run()
