import sys
import win32clipboard
import win32con
import win32print
import win32ui
import pyautogui
from pynput import mouse
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit, QLabel, QPushButton, QComboBox, QShortcut
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QKeySequence  # 正确的导入方式
import time
from pynput import mouse
import keyboard  # 确保导入keyboard库
class AutoPasteApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.is_selecting = False  # 选中即可粘贴开关
        self.is_enabled = False  # 打印开关
        self.thermal_printer_name = ""  # 存储选择的打印机名称
        # 清空剪贴板
        self.clear_clipboard()
        # 设置全局热键
        keyboard.add_hotkey('F9', self.toggleAutoPaste)  # 绑定F9自动填充
        keyboard.add_hotkey('F10', self.printContent)  # 绑定F10键到打印函数


    def initUI(self):
        self.setWindowTitle('自动填充文本框')

        layout = QVBoxLayout()

        self.textEdit = QTextEdit(self)
        layout.addWidget(self.textEdit)

        self.label = QLabel('选中文本将自动填入到此文本框', self)
        layout.addWidget(self.label)

        # 按钮1
        self.toggleButton = QPushButton('1启动自动填充(F9)', self)
        self.toggleButton.clicked.connect(self.toggleAutoPaste)
        layout.addWidget(self.toggleButton)

        # 按钮2
        self.printButton = QPushButton('2打印内容(F10)', self)
        self.printButton.clicked.connect(self.printContent)
        layout.addWidget(self.printButton)

        # 打印机选择下拉框
        self.printerComboBox = QComboBox(self)
        self.loadPrinters()
        layout.addWidget(self.printerComboBox)

        self.setLayout(layout)
        self.resize(400, 300)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)

        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.checkClipboard)
        self.timer.start(500)



    def loadPrinters(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        printer_names = [printer[2] for printer in printers]  # 获取打印机名称
        self.printerComboBox.addItems(printer_names)

    def checkClipboard(self):
        if not self.is_enabled:
            return

        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                current_text = self.textEdit.toPlainText()
                if current_text != data:
                    full_text = f"{data}"
                    self.textEdit.setPlainText(full_text)
            win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"读取剪贴板出错: {e}")

    def toggleAutoPaste(self):
        self.is_enabled = not self.is_enabled
        if self.is_enabled:
            self.toggleButton.setText('1停止自动填充(F9)')
            self.clear_clipboard()
        else:
            self.toggleButton.setText('1启动自动填充(F9)')

    def on_click(self, x, y, button, pressed):
        if self.is_enabled:
            if button == mouse.Button.left:
                if pressed:
                    self.is_selecting = True
                else:
                    self.is_selecting = False
                    self.copy_selection_to_clipboard()

    def copy_selection_to_clipboard(self):
        pyautogui.hotkey('ctrl', 'c')

    def printContent(self):
        content = self.textEdit.toPlainText().strip()
        self.thermal_printer_name = self.printerComboBox.currentText()  # 获取选择的打印机名称

        if content:
            hdc = None
            try:
                # 创建打印文档
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(self.thermal_printer_name)  # 使用选择的打印机名称
                hdc.StartDoc("Print Job")
                hdc.StartPage()

                # 设置纸张参数
                paper_size = hdc.GetDeviceCaps(win32con.HORZRES), hdc.GetDeviceCaps(win32con.VERTRES)
                left_margin = int(paper_size[0] * 0.1)  # 左边距10%
                top_margin = int(paper_size[1] * 0.1)  # 上边距10%
                line_height = int(paper_size[1] * 0.05)  # 每行高度
                line_height=line_height//2
                # 获取当前时间 + 打印内容
                current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
                text_to_print = f"{current_time} : {content}"

                # 设置字体
                font = win32ui.CreateFont({
                    "name": "Arial",
                    "height": line_height,
                    "weight": 1,
                })
                hdc.SelectObject(font)

                # 获取字体的宽度
                char_width = hdc.GetTextExtent("W")[0]  # 用字符 "W" 估算字符宽度
                row_weight = (paper_size[0] - left_margin * 2) // char_width  # 可以容纳的字符个数
                # row_weight=int(row_weight*1.2)
                current_y = top_margin  # 当前纵坐标
                for line in text_to_print.splitlines():
                    while len(line) > row_weight:
                        cur = line[:row_weight]
                        # 找到最后一个空格的位置，避免单词中断
                        last_space = cur.rfind(" ")
                        if last_space != -1:
                            cur = line[:last_space]
                            line = line[last_space + 1:]  # 更新行内容为剩余部分
                        else:
                            line = line[row_weight:]  # 若无空格则强制截断

                        hdc.TextOut(left_margin, current_y, cur)
                        current_y += line_height

                    # 打印当前行剩余的部分
                    if line:
                        hdc.TextOut(left_margin, current_y, line)
                        current_y += line_height  # 移动到下一行

                hdc.EndPage()
                hdc.EndDoc()
            except Exception as e:
                print(f"打印过程中发生错误: {e}")
            finally:
                if hdc:
                    hdc.DeleteDC()  # 确保在结束时删除设备上下文
        else:
            print("文本框为空或未更新，无法打印。")

    def clear_clipboard(self):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"清空剪贴板出错: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AutoPasteApp()
    ex.show()
    sys.exit(app.exec_())