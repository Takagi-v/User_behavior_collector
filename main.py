import os
import sys
import re
import time
import csv
import uuid
import datetime
import queue
import threading

# 尝试导入tkinter，兼容不同版本Python
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    # Python 2.x
    import Tkinter as tk
    import ttk
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox

# 尝试导入外部模块，如果导入失败给出清晰的错误信息和解决方案
missing_modules = []

try:
    import pygetwindow as gw
except ImportError:
    missing_modules.append(("PyGetWindow", "pip install PyGetWindow"))

try:
    import pyautogui
except ImportError:
    missing_modules.append(("PyAutoGUI", "pip install PyAutoGUI"))

try:
    import keyboard
except ImportError:
    missing_modules.append(("keyboard", "pip install keyboard"))

try:
    import mouse
except ImportError:
    missing_modules.append(("mouse", "pip install mouse"))

try:
    import pyperclip
except ImportError:
    missing_modules.append(("pyperclip", "pip install pyperclip"))

try:
    from PIL import ImageGrab, Image
except ImportError:
    missing_modules.append(("Pillow", "pip install Pillow"))

try:
    import psutil
except ImportError:
    missing_modules.append(("psutil", "pip install psutil"))
    psutil = None

# 尝试导入Win32 API模块，用于获取窗口进程信息
try:
    import win32gui
    import win32process
    has_win32 = True
except ImportError:
    missing_modules.append(("pywin32", "pip install pywin32"))
    has_win32 = False

# 如果有缺失模块，显示错误对话框
if missing_modules:
    print("\nERROR: Missing required modules!")
    print("Please install the following modules:")
    
    error_msg = "缺少以下必要模块，无法启动程序：\n\n"
    
    for module, cmd in missing_modules:
        print(f"  {cmd}")
        error_msg += f"• {module}\n"
    
    error_msg += "\n您可以通过以下方式安装：\n"
    error_msg += "1. 运行 setup_once.bat 脚本\n"
    error_msg += "2. 或者使用命令行手动安装：\n"
    
    for _, cmd in missing_modules:
        error_msg += f"   {cmd}\n"
    
    error_msg += "\n如果是在虚拟环境中运行，请确保使用：\n"
    error_msg += "venv\\Scripts\\pip.exe install [模块名]"

    # 使用tkinter显示错误消息（如果可用）
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("依赖错误", error_msg)
        root.destroy()
    except:
        print(error_msg)
    
    sys.exit(1)

class UserBehaviorCollector(tk.Tk):
    """
    用户行为数据采集工具
    监听键盘操作、鼠标操作及窗口切换事件
    """
    
    # 操作类型映射
    OPERATION_MAPPING = {
        "剪贴": ["键盘-组合键：Ctrl+X", "鼠标-操作：剪贴", "剪贴板操作：剪贴"],
        "复制": ["键盘-组合键：Ctrl+C", "键盘-组合键：Ctrl+Insert", "鼠标-右键菜单：复制", "剪贴板操作：复制"],
        "粘贴": ["键盘-组合键：Ctrl+V", "键盘-组合键：Shift+Insert", "鼠标-右键菜单：粘贴", "剪贴板操作：粘贴"],
        "删除": ["键盘-特殊键：Delete", "键盘-特殊键：Backspace"],
        "查看": ["截图：窗口切换", "窗口-状态：最小化", "窗口-状态：最大化", "窗口-状态：关闭",
                "键盘-特殊键：↑", "键盘-特殊键：↓", "键盘-特殊键：←", "键盘-特殊键：→",
                "键盘-特殊键：PageUp", "键盘-特殊键：PageDown", "鼠标-滚轮：向上", "鼠标-滚轮：向下",
                "鼠标-拖拽：滚动条"],
        "输入": ["键盘-输入：*", "键盘-特殊键：Space", "键盘-特殊键：Enter"],
        "点击": ["鼠标-单击：左键", "鼠标-单击：右键", "鼠标-双击：左键", "鼠标-双击：右键"],
        "其他": ["截图：定时截图", "键盘-组合键：Ctrl+A", "键盘-组合键：Ctrl+Z", "鼠标-操作：右键点击"]
    }
    
    def __init__(self):
        super().__init__()
        self.title("用户行为数据采集工具")
        self.geometry("500x400")
        self.resizable(False, False)
        
        # 初始化变量
        self.mac_address = self.get_mac_address()
        self.student_id = tk.StringVar()
        self.name = tk.StringVar()
        self.storage_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        
        # 数据队列
        self.data_queue = queue.Queue()
        
        # 监控状态
        self.monitoring = False
        self.paused = False
        
        # 线程列表
        self.threads = []
        
        # 日志缓存
        self.log_entries = []
        self.max_log_entries = 1000
        
        # 初始化界面
        self.init_ui()
    
    def get_mac_address(self):
        """获取MAC地址"""
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff)
                         for elements in range(0, 8*6, 8)][::-1])
        return mac
    
    def init_ui(self):
        """初始化界面"""
        # 创建主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # MAC地址显示
        ttk.Label(main_frame, text="MAC地址:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=self.mac_address).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 存储路径
        ttk.Label(main_frame, text="存储路径:").grid(row=1, column=0, sticky=tk.W, pady=5)
        path_frame = ttk.Frame(main_frame)
        path_frame.grid(row=1, column=1, sticky=tk.W+tk.E, pady=5)
        ttk.Entry(path_frame, textvariable=self.storage_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="浏览", command=self.browse_path).pack(side=tk.RIGHT)
        
        # 学号
        ttk.Label(main_frame, text="学号:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.student_id).grid(row=2, column=1, sticky=tk.W+tk.E, pady=5)
        
        # 姓名
        ttk.Label(main_frame, text="姓名:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.name).grid(row=3, column=1, sticky=tk.W+tk.E, pady=5)
        
        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        self.start_button = ttk.Button(button_frame, text="启用监控", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重置", command=self.reset).pack(side=tk.LEFT, padx=5)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W+tk.E+tk.N+tk.S, pady=5)
        
        # 创建滚动文本框
        self.log_text = tk.Text(log_frame, height=10, width=60)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 禁用日志文本框编辑
        self.log_text.config(state=tk.DISABLED)
        
        # 日志颜色标签
        self.log_text.tag_configure("error", foreground="red", font=("TkDefaultFont", 9, "bold"))
        
        # 设置窗口事件处理
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 设置列权重，使得控件在窗口调整大小时能够正确拉伸
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def browse_path(self):
        """选择存储路径"""
        path = filedialog.askdirectory()
        if path:
            self.storage_path.set(path)
    
    def reset(self):
        """重置所有输入并停止监控"""
        # 如果正在监控，先停止
        if self.monitoring:
            self.stop_monitoring()
        
        # 重置输入项
        self.student_id.set("")
        self.name.set("")
        self.storage_path.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        
        # 重置界面
        self.start_button.config(text="启用监控", command=self.start_monitoring)
        
        # 移除保存按钮
        if hasattr(self, 'save_button') and self.save_button.winfo_exists():
            self.save_button.destroy()
        
        self.log_activity("已重置所有设置")
    
    def validate_inputs(self):
        """验证输入的合法性"""
        # 检查学号和姓名是否为空
        if not self.student_id.get().strip():
            messagebox.showerror("错误", "学号不能为空!")
            return False
        
        if not self.name.get().strip():
            messagebox.showerror("错误", "姓名不能为空!")
            return False
        
        # 检查学号和姓名是否包含非法字符
        illegal_chars = r'[/\\:*?"<>|]'
        if re.search(illegal_chars, self.student_id.get()):
            messagebox.showerror("错误", "学号包含非法字符!")
            return False
        
        if re.search(illegal_chars, self.name.get()):
            messagebox.showerror("错误", "姓名包含非法字符!")
            return False
        
        # 检查存储路径是否存在
        storage_path = self.storage_path.get()
        if not os.path.exists(storage_path):
            try:
                os.makedirs(storage_path)
            except:
                messagebox.showerror("错误", "无法创建存储路径!")
                return False
        
        return True
    
    def start_monitoring(self):
        """启动监控"""
        if not self.validate_inputs():
            return
        
        if self.monitoring:
            return
        
        # 创建数据文件夹
        main_folder_path, csv_filepath, screenshots_folder = self.create_data_folder()
        if main_folder_path is None or csv_filepath is None or screenshots_folder is None:
            return
        
        self.main_folder_path = main_folder_path
        self.csv_filepath = csv_filepath
        self.screenshots_folder = screenshots_folder
        
        # 更新状态
        self.monitoring = True
        self.paused = False
        
        # 修改界面
        self.start_button.config(text="暂停", command=self.toggle_pause)
        
        # 获取父框架
        button_frame = self.start_button.master
        
        # 添加保存按钮 - 确保它位于正确的框架中
        if not hasattr(self, 'save_button') or not self.save_button.winfo_exists():
            self.save_button = ttk.Button(button_frame, text="保存", command=self.confirm_save)
            self.save_button.pack(side=tk.LEFT, padx=5)
            button_frame.update_idletasks()  # 强制更新布局
        
        # 调整窗口位置和大小
        self.geometry("350x350+0+0")
        self.attributes("-topmost", True)
        
        # 添加窗口透明度和展开/收缩功能
        self.bind("<Enter>", self.on_mouse_enter)
        self.bind("<Leave>", self.on_mouse_leave)
        
        # 记录开始时间
        self.start_time = datetime.datetime.now()
        
        # 启动各监听线程
        self.start_keyboard_listener()
        self.start_mouse_listener()
        self.start_window_listener()
        self.start_screenshot_timer()
        
        self.log_activity("监控已启动")
    
    def toggle_pause(self):
        """暂停/继续监控"""
        if not self.monitoring:
            return
        
        self.paused = not self.paused
        
        if self.paused:
            self.start_button.config(text="继续")
            self.log_activity("监控已暂停")
        else:
            self.start_button.config(text="暂停")
            self.log_activity("监控已继续")
    
    def confirm_save(self):
        """确认是否保存并退出"""
        result = messagebox.askyesno("确认", "是否停止程序并保存数据？")
        if result:
            self.stop_monitoring()
            # 打开存储文件夹
            try:
                os.startfile(self.main_folder_path)
            except:
                pass
            self.destroy()
    
    def stop_monitoring(self):
        """停止监控"""
        if not self.monitoring:
            return
        
        # 更新状态
        self.monitoring = False
        self.log_activity("正在停止监控...")
        
        # 停止所有线程
        for thread in self.threads:
            if thread.is_alive():
                # 这里只能标记线程停止，无法强制结束线程
                # 实际终止需要在线程循环中检测stop_flag
                thread.stop_flag = True
        
        # 等待线程停止（最多等待5秒）
        wait_time = 0
        while any(t.is_alive() for t in self.threads) and wait_time < 5:
            time.sleep(0.5)
            wait_time += 0.5
        
        # 释放键盘和鼠标钩子
        try:
            self.log_activity("正在释放键盘钩子...")
            keyboard.unhook_all()
        except Exception as e:
            self.log_activity(f"键盘钩子释放失败: {str(e)}", error=True)
        
        try:
            self.log_activity("正在释放鼠标钩子...")
            mouse.unhook_all()
        except Exception as e:
            self.log_activity(f"鼠标钩子释放失败: {str(e)}", error=True)
        
        # 清空线程列表
        self.threads.clear()
        
        self.log_activity("监控已停止")
        
        # 恢复��钮状态
        self.start_button.config(text="启用监控", command=self.start_monitoring)
        
        # 移除保存按钮（如果存在）
        if hasattr(self, 'save_button') and self.save_button.winfo_exists():
            self.save_button.destroy()
            delattr(self, 'save_button')
    
    def on_mouse_enter(self, event):
        """鼠标进入窗口"""
        self.attributes("-alpha", 1.0)  # 不透明度为100%
        # 调整窗口大小，确保有足够空间展示按钮和日志
        self.geometry("400x500+0+0")  # 增加窗口尺寸，确保内容完全显示
        
        # 确保界面元素全部可见
        self.update_idletasks()  # 强制更新界面布局
        
        # 确保按钮可见并置顶
        if hasattr(self, 'start_button'):
            self.start_button.lift()
        if hasattr(self, 'save_button'):
            self.save_button.lift()
    
    def on_mouse_leave(self, event):
        """鼠标离开窗口"""
        # 检查鼠标是否真的离开窗口（有时事件误触发）
        mouse_x, mouse_y = self.winfo_pointerxy()
        window_x = self.winfo_rootx()
        window_y = self.winfo_rooty()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        
        # 只有当鼠标真的在窗口外时才改变透明度
        if not (window_x <= mouse_x < window_x + window_width and 
                window_y <= mouse_y < window_y + window_height):
            self.attributes("-alpha", 0.3)  # 不透明度为30%
            self.geometry("350x150+0+0")  # 收缩窗口但保持足够宽度显示按钮
    
    def start_keyboard_listener(self):
        """启动键盘监听线程"""
        keyboard_thread = threading.Thread(target=self.keyboard_listener, daemon=True)
        keyboard_thread.stop_flag = False
        keyboard_thread.start()
        self.threads.append(keyboard_thread)
    
    def start_mouse_listener(self):
        """启动鼠标监听线程"""
        mouse_thread = threading.Thread(target=self.mouse_listener, daemon=True)
        mouse_thread.stop_flag = False
        mouse_thread.start()
        self.threads.append(mouse_thread)
    
    def start_window_listener(self):
        """启动窗口监听线程"""
        window_thread = threading.Thread(target=self.window_listener, daemon=True)
        window_thread.stop_flag = False
        window_thread.start()
        self.threads.append(window_thread)
    
    def start_screenshot_timer(self):
        """启动定时截图线程"""
        timer_thread = threading.Thread(target=self.screenshot_timer, daemon=True)
        timer_thread.stop_flag = False
        timer_thread.start()
        self.threads.append(timer_thread)
    
    def keyboard_listener(self):
        """键盘事件监听线程"""
        # 字符缓冲区
        buffer = []
        last_input_time = time.time()
        
        # 映射特殊键名
        special_key_map = {
            "enter": "回车",
            "backspace": "退格",
            "delete": "删除",
            "space": "空格",
            "up": "↑",
            "down": "↓",
            "left": "←",
            "right": "→",
            "page up": "PageUp",
            "page down": "PageDown"
        }
        
        # 组合键检测
        def on_hotkey(hotkey):
            nonlocal buffer
            
            # 如果暂停状态，不记录
            if self.paused:
                return
            
            # 清空缓冲区
            if buffer:
                input_text = "".join(buffer)
                operation_detail = f"键盘-输入：{input_text}"
                try:
                    window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                except:
                    window_title = "未知窗口"
                self.log_window_event(window_title, "NORMAL", operation_detail)
                buffer = []
            
            # 记录组合键
            operation_detail = f"键盘-组合键：{hotkey}"
            try:
                window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
            except:
                window_title = "未知窗口"
            self.log_window_event(window_title, "NORMAL", operation_detail)
        
        # 注册组合键
        try:
            keyboard.add_hotkey('ctrl+c', lambda: on_hotkey("Ctrl+C"))
            keyboard.add_hotkey('ctrl+insert', lambda: on_hotkey("Ctrl+Insert"))
            keyboard.add_hotkey('ctrl+v', lambda: on_hotkey("Ctrl+V"))
            keyboard.add_hotkey('shift+insert', lambda: on_hotkey("Shift+Insert"))
            keyboard.add_hotkey('ctrl+x', lambda: on_hotkey("Ctrl+X"))
            keyboard.add_hotkey('ctrl+z', lambda: on_hotkey("Ctrl+Z"))
            keyboard.add_hotkey('ctrl+a', lambda: on_hotkey("Ctrl+A"))
        except Exception as e:
            self.log_activity(f"组合键注册失败: {str(e)}", error=True)
        
        # 按键释放回调
        def on_key_release(e):
            nonlocal buffer, last_input_time
            
            # 如果暂停状态，不记录
            if self.paused:
                return
            
            key_name = e.name.lower() if hasattr(e, 'name') else ""
            
            # 特殊键处理
            if key_name in special_key_map:
                # 先处理缓冲区
                if buffer:
                    input_text = "".join(buffer)
                    operation_detail = f"键盘-输入：{input_text}"
                    try:
                        window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                    except:
                        window_title = "未知窗口"
                    self.log_window_event(window_title, "NORMAL", operation_detail)
                    buffer = []
                
                # 记录特殊键
                special_key = special_key_map[key_name]
                operation_detail = f"键盘-特殊键：{special_key}"
                try:
                    window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                except:
                    window_title = "未知窗口"
                self.log_window_event(window_title, "NORMAL", operation_detail)
                
            # 可打印字符处理
            elif len(key_name) == 1 or key_name in ['shift', 'ctrl', 'alt']:
                # ctrl, shift, alt 等修饰键不处理
                if key_name in ['shift', 'ctrl', 'alt']:
                    return
                
                # 添加到缓冲区
                buffer.append(key_name)
                last_input_time = time.time()
            
            # 检查组合键情况
            elif "+" in key_name:
                # 这是组合键，已在hotkey处理，这里不再处理
                pass
            
            # 处理其他键
            else:
                # 先处理缓冲区
                if buffer:
                    input_text = "".join(buffer)
                    operation_detail = f"键盘-输入：{input_text}"
                    try:
                        window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                    except:
                        window_title = "未知窗口"
                    self.log_window_event(window_title, "NORMAL", operation_detail)
                    buffer = []
                
                # 记录其他键
                operation_detail = f"键盘-特殊键：{key_name}"
                try:
                    window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                except:
                    window_title = "未知窗口"
                self.log_window_event(window_title, "NORMAL", operation_detail)
        
        # 注册按键释放事件
        try:
            keyboard.on_release(on_key_release)
        except Exception as e:
            self.log_activity(f"键盘监听失败: {str(e)}", error=True)
        
        # 主循环，定时检查缓冲区
        current_thread = threading.current_thread()
        while not getattr(current_thread, "stop_flag", False):
            # 如果缓冲区有内容且超过1秒没有新输入，则提交
            if buffer and (time.time() - last_input_time) > 1:
                if not self.paused:
                    input_text = "".join(buffer)
                    operation_detail = f"键盘-输入：{input_text}"
                    try:
                        window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                    except:
                        window_title = "未知窗口"
                    self.log_window_event(window_title, "NORMAL", operation_detail)
                buffer = []
            
            time.sleep(0.1)
    
    def mouse_listener(self):
        """鼠标事件监听线程"""
        # 鼠标点击事件
        def on_click(event=None, *args, **kwargs):
            # 如果暂停状态，不记录
            if self.paused:
                return
            
            # 处理不同调用方式
            # 1. 通过hook函数调用时，传递event对象
            # 2. 通过on_click函数调用时，传递x, y, button, pressed参数
            if event is None and len(args) >= 3:
                # on_click方式的调用
                x, y, button, pressed = args[0], args[1], args[2], args[3] if len(args) > 3 else False
                
                # 仅记录释放事件
                if pressed:
                    return
                
                # 确定按钮类型
                if hasattr(mouse, 'LEFT') and button == mouse.LEFT:
                    button_name = "左键"
                elif hasattr(mouse, 'RIGHT') and button == mouse.RIGHT:
                    button_name = "右键"
                else:
                    button_name = "左键" if str(button).lower() == "left" else "右键" if str(button).lower() == "right" else "中键"
            else:
                # hook方式的调用
                # 仅记录释放事件
                if hasattr(event, 'event_type') and event.event_type != 'up':
                    return
                
                # 确定鼠标按钮
                button_name = "左键"
                if hasattr(event, 'button'):
                    if event.button == 'right':
                        button_name = "右键"
                    elif event.button == 'middle':
                        button_name = "中键"
            
            # 记录事件
            operation_detail = f"鼠标-单击：{button_name}"
            
            try:
                window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
            except:
                window_title = "未知窗口"
                
            self.log_window_event(window_title, "NORMAL", operation_detail)
            return False  # 不拦截事件
        
        # 滚轮事件监听函数
        def on_wheel(event=None, *args, **kwargs):
            # 如果暂停状态，不记录
            if self.paused:
                return
            
            # 处理不同调用方式
            if event is None and len(args) >= 3:
                # on_scroll方式的调用 (x, y, dx, dy)
                x, y, dx, dy = args[0], args[1], args[2], args[3] if len(args) > 3 else 0
                direction = "向上" if dy > 0 else "向下"
            else:
                # hook方式的调用
                direction = "向上" if hasattr(event, 'delta') and event.delta > 0 else "向下"
            
            # 记录事件
            operation_detail = f"鼠标-滚轮：{direction}"
            
            try:
                window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
            except:
                window_title = "未知窗口"
                
            self.log_window_event(window_title, "NORMAL", operation_detail)
            return False  # 不拦截事件
        
        # 通用鼠标事件处理函数
        def generic_mouse_hook(event):
            if hasattr(event, 'event_type'):
                if event.event_type in ['up', 'down']:
                    return on_click(event)
                elif event.event_type == 'wheel':
                    return on_wheel(event)
            return False
        
        # 尝试使用不同方法注册鼠标事件
        try:
            self.log_activity("正在注册鼠标钩子...")
            # 首选使用hook方法，兼容性更好
            mouse.hook(generic_mouse_hook)
            self.log_activity("成功注册鼠标通用钩子")
        except Exception as e:
            self.log_activity(f"通用钩子注册失败: {str(e)}", error=True)
            try:
                # 尝试分别注册点击和滚轮
                if hasattr(mouse, 'on_click'):
                    mouse.on_click(on_click)
                    self.log_activity("成功注册鼠标点击事件")
                
                if hasattr(mouse, 'on_scroll'):
                    mouse.on_scroll(on_wheel)
                    self.log_activity("成功注册鼠标滚轮事件")
                else:
                    # 备用滚轮监听方法
                    def mouse_wheel_alternative():
                        try:
                            import pyautogui
                            last_scroll = 0
                            
                            while not getattr(threading.current_thread(), "stop_flag", False):
                                # 尝试检测滚轮变化
                                for i in range(10):  # 每次检查10次
                                    if pyautogui._mouseScrolled:
                                        direction = "向上" if pyautogui._mouseScrollAmount > 0 else "向下"
                                        operation_detail = f"鼠标-滚轮：{direction}"
                                        
                                        try:
                                            window_title = gw.getActiveWindow().title if gw.getActiveWindow() else "未知窗口"
                                        except:
                                            window_title = "未知窗口"
                                            
                                        self.log_window_event(window_title, "NORMAL", operation_detail)
                                        pyautogui._mouseScrolled = False
                                    time.sleep(0.01)
                                time.sleep(0.1)
                        except:
                            self.log_activity("备用滚轮监听失败", error=True)
                    
                    # 创建备用滚轮监听线程
                    wheel_thread = threading.Thread(target=mouse_wheel_alternative, daemon=True)
                    wheel_thread.stop_flag = False
                    wheel_thread.start()
                    self.threads.append(wheel_thread)
                    self.log_activity("已启动备用滚轮监听")
            except Exception as e:
                self.log_activity(f"无法注册鼠标事件: {str(e)}", error=True)
        
        # 主循环，保持线程运行
        current_thread = threading.current_thread()
        while not getattr(current_thread, "stop_flag", False):
            time.sleep(0.1)
    
    def window_listener(self):
        """窗口事件监听线程"""
        last_window_title = None
        last_window_state = None
        
        current_thread = threading.current_thread()
        while not getattr(current_thread, "stop_flag", False):
            if self.paused:
                time.sleep(0.5)
                continue
            
            try:
                # 获取当前活动窗口
                current_window = gw.getActiveWindow()
                
                if current_window:
                    current_title = current_window.title
                    
                    # 检测窗口状态
                    try:
                        if current_window.isMaximized:
                            current_state = "MAXIMIZED"
                        elif current_window.isMinimized:
                            current_state = "MINIMIZED"
                        else:
                            current_state = "NORMAL"
                    except:
                        current_state = "NORMAL"  # 默认状态
                    
                    # 窗口切换检测
                    if last_window_title is None or last_window_title != current_title:
                        # 窗口切换，记录事件并截图
                        try:
                            self.take_screenshot("窗口切换")
                        except Exception as e:
                            self.log_activity(f"窗口切换截图失败: {str(e)}", error=True)
                        
                        # 记录窗口状态变化
                        last_window_title = current_title
                    
                    # 窗口状态变化检测
                    elif last_window_state is not None and last_window_state != current_state:
                        # 窗口状态变化，记录事件
                        operation_detail = f"窗口-状态：{current_state}"
                        self.log_window_event(current_title, current_state, operation_detail)
                    
                    last_window_state = current_state
            except Exception as e:
                self.log_activity(f"窗口监听错误: {str(e)}", error=True)
                # 重置状态，避免连续错误
                last_window_title = None
                last_window_state = None
            
            # 每隔0.5秒检测一次
            time.sleep(0.5)
    
    def screenshot_timer(self):
        """定时截图线程"""
        current_thread = threading.current_thread()
        while not getattr(current_thread, "stop_flag", False):
            if not self.paused:
                # 每60秒截图一次
                self.take_screenshot("定时截图")
            
            # 睡眠60秒
            for _ in range(600):  # 60秒，每0.1秒检查一次停止标志
                if getattr(current_thread, "stop_flag", False):
                    break
                time.sleep(0.1)

    def create_data_folder(self):
        """创建数据存储文件夹"""
        # 获取当前时间
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        
        # 构建文件夹名称: 学号_姓名_开始日期_开始时间
        folder_name = f"{self.student_id.get()}_{self.name.get()}_{now.strftime('%y%m%d')}_{now.strftime('%H%M%S')}"
        
        # 创建主文件夹路径
        main_folder_path = os.path.join(self.storage_path.get(), folder_name)
        
        # 创建主文件夹
        try:
            os.makedirs(main_folder_path, exist_ok=True)
        except:
            messagebox.showerror("错误", "无法创建存储文件夹!")
            return None, None, None
        
        # 创建截图子文件夹
        screenshots_folder = os.path.join(main_folder_path, "screenshots")
        try:
            os.makedirs(screenshots_folder, exist_ok=True)
        except:
            messagebox.showerror("错误", "无法创建截图文件夹!")
            return None, None, None
        
        # 创建CSV文件
        csv_filename = f"{self.student_id.get()}_{self.name.get()}_{now.strftime('%Y%m%d')}_{now.strftime('%H%M%S')}.csv"
        csv_filepath = os.path.join(main_folder_path, csv_filename)
        
        # 创建CSV文件并写入表头
        try:
            with open(csv_filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['学号', '姓名', '时间戳', '操作类型', '操作详情', '窗口进程名', '窗口标签', '窗口状态', '剪贴板内容', '截图文件'])
        except:
            messagebox.showerror("错误", "无法创建CSV文件!")
            return None, None, None
        
        return main_folder_path, csv_filepath, screenshots_folder
        
    def take_screenshot(self, reason="窗口切换"):
        """捕获当前活动窗口的截图"""
        try:
            # 获取当前活动窗口
            active_window = None
            try:
                active_window = gw.getActiveWindow()
                if not active_window:
                    self.log_activity("无法获取活动窗口，使用全屏截图", error=False)
            except Exception as e:
                self.log_activity(f"获取活动窗口失败: {str(e)}，使用全屏截图", error=True)
            
            # 获取窗口标签
            window_title = "未知窗口"
            if active_window:
                try:
                    window_title = active_window.title
                except:
                    window_title = "无法获取窗口标题"
            
            # 清理窗口标签中的非法字符
            cleaned_title = re.sub(r'[\\/*?:"<>|]', '-', window_title)
            # 限制文件名长度
            cleaned_title = cleaned_title[:30]
            
            # 生成截图文件名: 学号_开始时间_毫秒_窗口标签.png
            now = datetime.datetime.now()
            timestamp = now.strftime("%H%M%S")
            milliseconds = now.microsecond // 1000
            
            screenshot_filename = f"{self.student_id.get()}_{timestamp}_{milliseconds:02d}_{cleaned_title}.png"
            screenshot_path = os.path.join(self.screenshots_folder, screenshot_filename)
            
            # 捕获截图
            try:
                screenshot = ImageGrab.grab()
                screenshot.save(screenshot_path)
            except Exception as e:
                self.log_activity(f"截图保存失败: {str(e)}", error=True)
                return None
            
            # 生成操作详情
            operation_detail = f"截图：{reason}"
            
            # 记录到CSV
            self.log_window_event(window_title, "NORMAL", operation_detail, screenshot_filename)
            
            # 记录到日志
            self.log_activity(f"【查看】 {operation_detail} {screenshot_filename}")
            
            return screenshot_filename
        except Exception as e:
            self.log_activity(f"截图过程出错: {str(e)}", error=True)
            return None
            
    def log_window_event(self, window_title, window_state, operation_detail, screenshot_filename=None):
        """记录窗口事件到CSV"""
        try:
            # 获取当前活动窗口的进程名
            process_name = "未知"
            
            # 使用不同方法尝试获取进程名
            try:
                active_window = gw.getActiveWindow()
                
                # 方法1: 如果可用，使用Win32 API获取PID
                if has_win32 and psutil is not None:
                    try:
                        hwnd = win32gui.GetForegroundWindow()
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(pid)
                        process_name = process.name()
                    except Exception as e:
                        # 仅调试日志
                        # self.log_activity(f"获取进程名方法1失败: {str(e)}", error=False)
                        pass
                
                # 方法2: 如果方法1失败，尝试使用pygetwindow自带的属性
                if process_name == "未知" and active_window and psutil is not None:
                    try:
                        if hasattr(active_window, '_hWnd'):
                            pid = active_window._hWnd
                            try:
                                # 尝试直接将_hWnd作为PID使用
                                process = psutil.Process(pid)
                                process_name = process.name()
                            except:
                                pass
                    except:
                        pass
            except:
                pass
            
            # 从窗口标题中提取核心内容（针对浏览器）
            window_core_title = window_title
            if any(browser in window_title.lower() for browser in ["chrome", "firefox", "edge", "opera", "safari"]):
                # 尝试使用常见的分隔符分割标题
                for separator in [" - ", " | ", " — "]:
                    if separator in window_title:
                        parts = window_title.split(separator)
                        if len(parts) > 1:
                            window_core_title = parts[1].strip()
                            break
            
            # 准备CSV数据
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 获取剪贴板内容
            clipboard_content = ""
            if any(keyword in operation_detail for keyword in ["复制", "剪贴", "粘贴"]):
                try:
                    clip_text = pyperclip.paste()
                    if clip_text:
                        # 格式化剪贴板内容: 【字数】内容（来源：窗口标签）
                        text_len = len(clip_text)
                        if text_len > 200:
                            # 如果超过200字，只显示前50个和后50个字符
                            formatted_content = f"【{text_len}】{clip_text[:50]}......{clip_text[-50:]}（来源：{window_core_title}）"
                        else:
                            formatted_content = f"【{text_len}】{clip_text}（来源：{window_core_title}）"
                        clipboard_content = formatted_content
                except:
                    clipboard_content = "无法获取剪贴板内容"
            
            # 操作类型映射
            operation_type = "其他"  # 默认操作类型
            for op_type, patterns in self.OPERATION_MAPPING.items():
                if any(pattern in operation_detail for pattern in patterns) or any(pattern.startswith(operation_detail) for pattern in patterns):
                    operation_type = op_type
                    break
            
            # 准备数据行
            data_row = [
                self.student_id.get(),
                self.name.get(),
                timestamp,
                operation_type,
                operation_detail,
                process_name,
                window_core_title,
                window_state,
                clipboard_content,
                screenshot_filename or ""
            ]
            
            # 写入CSV
            self.write_to_csv(data_row)
            
        except Exception as e:
            self.log_activity(f"记录窗口事件失败: {str(e)}", error=True)

    def log_activity(self, log_text, error=False):
        """记录活动到界面日志"""
        # 添加时间戳
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{now}] {log_text}"
        
        # 添加到日志列表
        self.log_entries.append(log_entry)
        
        # 如果日志超过最大数量，移除最早的记录
        if len(self.log_entries) > self.max_log_entries:
            self.log_entries.pop(0)
        
        # 更新界面日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        
        # 如果日志条目很多，只显示最新的10条
        entries_to_show = self.log_entries[-10:] if len(self.log_entries) > 10 else self.log_entries
        
        for entry in entries_to_show:
            if error and entry == log_entry:
                self.log_text.insert(tk.END, entry + "\n", "error")
            else:
                self.log_text.insert(tk.END, entry + "\n")
        
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)  # 滚动到最新内容
        
    def write_to_csv(self, data):
        """写入数据到CSV文件"""
        try:
            with open(self.csv_filepath, 'a', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(data)
        except Exception as e:
            self.log_activity(f"错误：无法写入CSV文件 - {str(e)}", error=True)

    def on_close(self):
        """窗口关闭处理"""
        if self.monitoring:
            result = messagebox.askyesno("确认", "监控正在进行中，确定要退出程序吗？")
            if result:
                self.stop_monitoring()
                self.destroy()
        else:
            self.destroy()

if __name__ == "__main__":
    app = UserBehaviorCollector()
    app.mainloop()