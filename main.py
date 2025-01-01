from typing import Any, List
import json
import os
import sys
import time

import numpy as np
import win32api
import win32con
import win32gui
from PIL import ImageGrab
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit, \
    QSizePolicy, QLabel
from PySide6.QtGui import QIcon
from fuzzywuzzy import process
from paddleocr import PaddleOCR


class WindowHandler:
    def __init__(self):
        self.window = None
        self.window_title = "咸鱼之王"
        self.find_window()

    def find_window(self):
        def callback(hwnd, extra):
            if self.window_title in win32gui.GetWindowText(hwnd):
                self.window = hwnd
                return False
            return True

        win32gui.EnumWindows(callback, None)
        if not self.window:
            raise Exception(f"未找到{self.window_title}窗口")

    def capture_screenshot_ext(self, left, top, right, bottom):
        try:
            self.find_window()
            try:
                if self.window and win32gui.IsWindow(self.window):
                    placement = win32gui.GetWindowPlacement(self.window)
                    if placement[1] == win32con.SW_SHOWMINIMIZED:
                        win32gui.ShowWindow(self.window, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(self.window)
                    time.sleep(0.1)
            except:
                pass

            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            return np.array(screenshot)

        except Exception as e:
            print(f"截图失败: {e}")
            return np.zeros((bottom - top, right - left, 3), dtype=np.uint8)


class WinOperator:
    def __init__(self, window):
        self.window = window

    def click(self, x, y):
        """
        在指定坐标发送点击消息，不移动鼠标
        """
        try:
            # 确保窗口是激活的
            if self.window and win32gui.IsWindow(self.window):
                # 将坐标转换为窗口客户区坐标
                left, top, right, bottom = win32gui.GetWindowRect(self.window)
                x = x - left
                y = y - top

                # 将坐标打包成LPARAM
                lParam = win32api.MAKELONG(x, y)

                # 发送鼠标消息
                win32gui.SendMessage(self.window, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
                time.sleep(0.1)
                win32gui.SendMessage(self.window, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)
                # print(f"已发送点击消息到坐标: ({x}, {y})")
                return True
            else:
                print("无效的窗口句柄")
                return False

        except Exception as e:
            print(f"点击操作失败: {e}")
            return False


class Ocr:
    def __init__(self) -> None:
        self.ocr = PaddleOCR(show_log=False)
        self.data = None  # 存储OCR识别结果

    def do_ocr_ext(self, img_data, simple=False) -> List:
        data = self.ocr.ocr(img_data, cls=False)[0]
        if simple: return self.get_all_text(data)
        self.data = data
        return data

    def get_all_text(self, data: List[List[Any]] = None, position=False):
        """
        返回所有文本及其位置

        参数:
        data (List[List[Any]]): OCR识别结果的数据。

        返回:
        None
        """
        data = data if data else self.data
        res = []
        if data is None: return res
        for item in data:
            text = str(item[1][0])  # 确保 text 是字符串类型
            points = item[0]
            res.append((text, points) if position else text)
        return res


class ConsoleOutput:
    def __init__(self, text_edit):
        self.text_edit = text_edit

    def write(self, text):
        # 使用 Signal 在主线程中更新 UI
        if hasattr(self.text_edit, 'append_text'):
            self.text_edit.append_text.emit(text.rstrip())
        else:
            self.text_edit.append(text.rstrip())

    def flush(self):
        pass


class SafeTextEdit(QTextEdit):
    append_text = Signal(str)

    def __init__(self):
        super().__init__()
        self.append_text.connect(self.append)
        self.setReadOnly(True)
        # 设置最大行数限制，避免内存占用过大
        self.document().setMaximumBlockCount(1000)


class WorkerThread(QThread):
    finished = Signal()
    error = Signal(str)

    def __init__(self, worker):
        super().__init__()
        self.worker = worker
        self.is_running = True

    def run(self):
        try:
            self.worker.run()
        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()

    def stop(self):
        if self.worker:
            self.worker.stop()
        self.is_running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("臭咸鱼答题助手v1.0")
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        # 设置固定尺寸
        self.setFixedSize(300, 400)

        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(10, 10, 10, 10)  # 设置边距

        # 创建线程安全的控制台输出显示区域
        self.console_output = SafeTextEdit()
        layout.addWidget(self.console_output)

        # 创建按钮容器和水平布局
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 0, 0, 0)  # 移除按钮容器边距
        button_layout.setSpacing(10)  # 设置按钮之间的间距

        # 创建按钮
        self.start_button = QPushButton("开始答题")
        self.stop_button = QPushButton("停止答题")
        self.stop_button.setEnabled(False)

        # 设置按钮样式和大小
        button_style = """
            QPushButton {
                min-height: 40px;
                font-size: 14px;
            }
        """
        self.start_button.setStyleSheet(button_style)
        self.stop_button.setStyleSheet(button_style)

        # 设置按钮大小策略
        self.start_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.stop_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # 添加按钮到水平布局
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)

        # 添加按钮容器到主布局
        layout.addWidget(button_container)

        # 创建超链接标签
        link_label = QLabel()
        link_label.setText('<a href="https://docs.qq.com/doc/DS1RzUFhnaFZoYXV3">帮助文档及交流群</a>')
        link_label.setAlignment(Qt.AlignLeft)  # 改为左对齐
        link_label.setContentsMargins(5, 0, 0, 0)  # 添加左边距，使其不贴边
        link_label.setOpenExternalLinks(True)  # 允许打开外部链接
        layout.addWidget(link_label)

        # 连接按钮信号
        self.start_button.clicked.connect(self.start_answering)
        self.stop_button.clicked.connect(self.stop_answering)

        # 重定向标准输出到文本框
        sys.stdout = ConsoleOutput(self.console_output)

        # 初始化工作线程
        self.worker = None
        self.thread = None

    def start_answering(self):
        if not self.thread or not self.thread.isRunning():
            try:
                self.worker = MainWorker()
                self.thread = WorkerThread(self.worker)
                self.thread.finished.connect(self.on_finished)
                self.thread.error.connect(self.on_error)

                self.thread.start()
                self.start_button.setEnabled(False)
                self.stop_button.setEnabled(True)
            except Exception as e:
                self.console_output.append(f"启动失败: {str(e)}")

    def stop_answering(self):
        if self.thread and self.thread.isRunning():
            try:
                self.thread.stop()
                # 减少等待时间
                self.thread.wait(1000)  # 等待最多1秒
                if self.thread.isRunning():
                    self.thread.terminate()
                self.on_finished()
            except Exception as e:
                self.console_output.append(f"停止失败: {str(e)}")

    def on_finished(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def on_error(self, error_msg):
        self.console_output.append(f"错误: {error_msg}")
        self.on_finished()

    def closeEvent(self, event):
        # 窗口关闭时确保线程正确停止
        if self.thread and self.thread.isRunning():
            self.stop_answering()
        event.accept()


def get_window_rect(window_title):
    """
    获取指定标题窗口的位置和大小
    Args:
        window_title: 窗口标题（部分匹配即可）
    Returns:
        tuple: (left, top, right, bottom) 窗口的坐标，未找到则返回 None
    """

    def callback(hwnd, extra):
        if window_title in win32gui.GetWindowText(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            extra.append(rect)

    rects = []
    win32gui.EnumWindows(callback, rects)

    if rects:
        return rects[0]
    return None


def get_recognition_area():
    """
    获取并计算识别区域的坐标
    Returns:
        tuple: (left, top, right, bottom) 识别区域的坐标，未找到窗口则返回 None
    """
    window_rect = get_window_rect("咸鱼之王")
    if window_rect:
        left, top, right, bottom = window_rect
        window_height = bottom - top
        window_width = right - left

        recognition_left = int(left + window_width * 0.02)
        recognition_right = int(left + window_width * 0.75)
        recognition_top = int(top + window_height * 0.13)
        recognition_bottom = int(top + window_height * 0.25)

        return (recognition_left, recognition_top, recognition_right, recognition_bottom)
    return None


def get_confirm_button_area():
    """
    获取"确定"按钮的识别区域
    Returns:
        tuple: (left, top, right, bottom) 按钮区域的坐标，未找到窗口则返回 None
    """
    window_rect = get_window_rect("咸鱼之王")
    if window_rect:
        left, top, right, bottom = window_rect
        window_height = bottom - top
        window_width = right - left

        button_top = int(top + window_height * 0.785)
        button_bottom = int(top + window_height * 0.87)
        button_left = int(left + window_width * 0.3)
        button_right = int(left + window_width * 0.7)

        return (button_left, button_top, button_right, button_bottom)
    return None


def find_best_match(properties, query):
    names = [prop['q'] for prop in properties]
    best_match = process.extractOne(query, names)
    if best_match:
        similarity = best_match[1]
        if similarity < 30:
            return None
        best_name = best_match[0]
        for prop in properties:
            if prop['q'] == best_name:
                return prop
    return None


def parse_json_lines(file_path):
    # 返回一个json列表
    json_list = []
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            try:
                json_data = json.loads(line)
                json_list.append(json_data)
                # yield json_data
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON on line: {line}")
                print(e)
        return json_list


def check_confirm_button(handler, ocr, operator):
    """检查按钮区域的文字"""
    button_area = get_confirm_button_area()
    if not button_area:
        return ""

    left, top, right, bottom = button_area
    screenshot = handler.capture_screenshot_ext(left, top, right, bottom)

    if screenshot.size == 0:
        return ""

    text = ''.join(ocr.do_ocr_ext(screenshot, simple=True))

    if "开始答题" in text:
        return "开始答题"
    elif "确定" in text:
        return "确定"
    return ""


class MainWorker:
    def __init__(self):
        self.is_running = True

    def stop(self):
        self.is_running = False

    def check_stop(self):
        """检查是否应该停止，并适当休眠"""
        if not self.is_running:
            return True
        # 使用更短的休眠时间，多次检查是否停止
        for _ in range(10):
            if not self.is_running:
                return True
            time.sleep(0.1)
        return False

    def run(self):
        handler = WindowHandler()
        operator = WinOperator(handler.window)
        ocr = Ocr()

        def click_answer(answer_text):
            window_rect = get_window_rect("咸鱼之王")
            if not window_rect:
                return False

            left, top, right, bottom = window_rect
            window_width = right - left
            window_height = bottom - top
            click_y = top + int(window_height * 0.8)

            if answer_text.strip().upper() == 'A':
                click_x = left + int(window_width * 0.3)
            elif answer_text.strip().upper() == 'B':
                click_x = left + int(window_width * 0.7)
            else:
                return False

            return operator.click(click_x, click_y)

        # 加载问答数据
        results = []
        for root, dirs, files in os.walk("data"):
            for file in files:
                result = parse_json_lines(os.path.join(root, file))
                results.extend(result)

        start_delay = 3
        answer_delay = 3

        recognition_area = get_recognition_area()
        if not recognition_area:
            print("未找到游戏窗口")
            return

        answering = False
        start_button_clicked = False

        # 检查是否在正确界面的计时器
        check_start_time = time.time()
        start_check_timeout = 0  # 0秒超时
        start_button_found = False

        while self.is_running:
            try:
                button_text = check_confirm_button(handler, ocr, operator)

                # 检查是否找到开始答题按钮
                if button_text == "开始答题":
                    start_button_found = True

                # 如果3秒内没有检测到开始答题按钮，提示并退出
                if not start_button_found and time.time() - check_start_time > start_check_timeout:
                    print("\n请在咸鱼大冲关界面启动")
                    break

                if button_text == "确定":
                    print("\n检测到确定按钮，停止运行")
                    break

                elif button_text == "开始答题":
                    if start_button_clicked:
                        print("\n答题按钮点击无响应，停止运行")
                        break

                    if not answering:
                        print("\n开始答题")
                        button_area = get_confirm_button_area()
                        if button_area:
                            left, top, right, bottom = button_area
                            center_x = (left + right) // 2
                            center_y = (top + bottom) // 2
                            operator.click(center_x, center_y)

                        answering = True
                        start_button_clicked = True
                        print(f"等待 {start_delay} 秒...")
                        for _ in range(int(start_delay * 10)):
                            if not self.is_running:
                                return
                            time.sleep(0.1)

                if answering:
                    recognition_area = get_recognition_area()
                    if not recognition_area:
                        if self.check_stop():
                            break
                        continue

                    screenshot_data = handler.capture_screenshot_ext(*recognition_area)
                    if screenshot_data.size == 0 or len(
                            question := ''.join(ocr.do_ocr_ext(screenshot_data, simple=True))) == 0:
                        if self.check_stop():
                            break
                        continue

                    if answer := find_best_match(results, question):
                        print(f"\n{answer['q']} ---> {answer['ans']}")
                        click_answer(answer['ans'])
                        for _ in range(int(answer_delay * 10)):
                            if not self.is_running:
                                return
                            time.sleep(0.1)
                    else:
                        if self.check_stop():
                            break
                        time.sleep(0.5)
                        continue

                    if self.check_stop():
                        break
                    continue

                else:
                    if self.check_stop():
                        break
                    time.sleep(0.5)
                    continue

            except Exception as e:
                print(f"错误: {e}")
                if not self.is_running or self.check_stop():
                    break


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
