"""
弹窗工具模块 - 统一处理所有弹窗的图层显示问题
"""
import tkinter as tk
from tkinter import ttk

def create_topmost_messagebox(parent, title, message, msg_type="info", buttons="ok"):
    """
    创建始终置顶的消息框
    
    Args:
        parent: 父窗口
        title: 标题
        message: 消息内容
        msg_type: 消息类型 ("info", "warning", "error", "question")
        buttons: 按钮类型 ("ok", "yesno", "yesnocancel")
    
    Returns:
        对于yesno/yesnocancel类型返回True/False/None，其他返回None
    """
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    dialog.focus_force()
    dialog.resizable(False, False)
    
    # 立即隐藏窗口，避免在左上角显示
    dialog.withdraw()
    
    # 设置图标和颜色
    icon_colors = {
        "info": "#1E90FF",
        "warning": "#FF8C00", 
        "error": "#DC143C",
        "question": "#32CD32"
    }
    
    icons = {
        "info": "ℹ",
        "warning": "⚠",
        "error": "✖",
        "question": "?"
    }
    
    # 主框架
    main_frame = ttk.Frame(dialog, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 图标和消息框架
    content_frame = ttk.Frame(main_frame)
    content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
    
    # 图标
    icon_label = tk.Label(content_frame, 
                         text=icons.get(msg_type, "ℹ"),
                         font=('Microsoft YaHei UI', 24),
                         fg=icon_colors.get(msg_type, "#1E90FF"),
                         bg='white')
    icon_label.pack(side=tk.LEFT, padx=(0, 15))
    
    # 消息文本
    msg_label = tk.Label(content_frame,
                        text=message,
                        font=('Microsoft YaHei UI', 10),
                        bg='white',
                        wraplength=400,
                        justify=tk.LEFT)
    msg_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 按钮框架
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X)
    
    result = None
    
    def on_yes():
        nonlocal result
        result = True
        dialog.destroy()
    
    def on_no():
        nonlocal result
        result = False
        dialog.destroy()
    
    def on_cancel():
        nonlocal result
        result = None
        dialog.destroy()
    
    def on_ok():
        nonlocal result
        result = None
        dialog.destroy()
    
    # 根据按钮类型创建按钮
    if buttons == "yesno":
        yes_btn = ttk.Button(button_frame, text="是", command=on_yes)
        yes_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        no_btn = ttk.Button(button_frame, text="否", command=on_no)
        no_btn.pack(side=tk.RIGHT)
        
        # 默认焦点在是按钮
        yes_btn.focus_set()
        
    elif buttons == "yesnocancel":
        cancel_btn = ttk.Button(button_frame, text="取消", command=on_cancel)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        no_btn = ttk.Button(button_frame, text="否", command=on_no)
        no_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        yes_btn = ttk.Button(button_frame, text="是", command=on_yes)
        yes_btn.pack(side=tk.RIGHT)
        
        # 默认焦点在是按钮
        yes_btn.focus_set()
        
    else:  # "ok"
        ok_btn = ttk.Button(button_frame, text="确定", command=on_ok)
        ok_btn.pack(side=tk.RIGHT)
        ok_btn.focus_set()
    
    # 居中显示 - 先更新布局但窗口仍然隐藏
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # 设置好位置后再显示窗口，避免移动效果
    dialog.deiconify()
    
    # 绑定ESC键关闭
    dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    # 等待用户操作
    dialog.wait_window()
    
    return result

def showinfo(title, message, parent=None):
    """显示信息对话框"""
    return create_topmost_messagebox(parent, title, message, "info", "ok")

def showwarning(title, message, parent=None):
    """显示警告对话框"""
    return create_topmost_messagebox(parent, title, message, "warning", "ok")

def showerror(title, message, parent=None):
    """显示错误对话框"""
    return create_topmost_messagebox(parent, title, message, "error", "ok")

def askyesno(title, message, parent=None):
    """显示是否对话框"""
    return create_topmost_messagebox(parent, title, message, "question", "yesno")

def askyesnocancel(title, message, parent=None):
    """显示是否取消对话框"""
    return create_topmost_messagebox(parent, title, message, "question", "yesnocancel")

def askquestion(title, message, parent=None):
    """显示问题对话框（兼容性）"""
    result = create_topmost_messagebox(parent, title, message, "question", "yesno")
    return "yes" if result else "no"