import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinterdnd2 as tkdnd
import popup_utils
# PIL 延迟导入 - 只在需要时才加载
# from PIL import Image, ImageTk
from smart_upscale_plan_dialog import show_smart_upscale_plan_dialog
import os
import sys
import shutil
import json
import threading
from pathlib import Path
import re
import warnings
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
import time
import pickle
import hashlib

# PIL 延迟导入全局变量
_PIL_Image = None
_PIL_ImageTk = None

def _get_PIL_Image():
    """延迟导入 PIL.Image"""
    global _PIL_Image
    if _PIL_Image is None:
        from PIL import Image
        _PIL_Image = Image
        # 设置 PIL 配置
        warnings.filterwarnings("ignore", category=Image.DecompressionBombWarning)
        Image.MAX_IMAGE_PIXELS = None
    return _PIL_Image

def _get_PIL_ImageTk():
    """延迟导入 PIL.ImageTk"""
    global _PIL_ImageTk
    if _PIL_ImageTk is None:
        from PIL import ImageTk
        _PIL_ImageTk = ImageTk
    return _PIL_ImageTk

# --- 调试输出过滤器 ---
DEBUG = False  # 切换为 True 可输出全部调试信息
# 需要过滤的 emoji（不包含❌，保留错误提示）
# 扩展过滤列表，加入 🧩 和 📊 以及常见调试前缀
_FILTER_EMOJIS = {'🔍','📏','✅','🎯','🏆','📦','🔧','ℹ️','🚫','🧩','📊'}
# 保留输出的关键字
_PROTECT_KEYWORDS = ('合并A列', '合并D列', 'A列和D列', '合并A列和D列', 'Exception', '错误', '失败')

import builtins as _builtins
_original_print = _builtins.print

def _filtered_print(*args, **kwargs):
    """过滤掉无用调试信息，保留关键日志/异常"""
    if not DEBUG and args and isinstance(args[0], str):
        first = str(args[0])
        # 如果包含需保护关键字，则直接输出
        if any(key in first for key in _PROTECT_KEYWORDS):
            return _original_print(*args, **kwargs)
        # 否则若包含过滤emoji，则跳过
        if any(em in first for em in _FILTER_EMOJIS):
            return
    _original_print(*args, **kwargs)

_builtins.print = _filtered_print

# 全局替换messagebox为popup_utils
messagebox.showinfo = popup_utils.showinfo
messagebox.showwarning = popup_utils.showwarning
messagebox.showerror = popup_utils.showerror

# 禁用PIL警告（延迟到实际使用时设置）
# warnings.filterwarnings("ignore", category=Image.DecompressionBombWarning)
# 提高PIL图像大小限制（延迟到实际使用时设置）
# Image.MAX_IMAGE_PIXELS = None

# 支持的图像格式
SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp'}

def get_app_directory():
    """获取应用程序根目录"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        return os.path.dirname(sys.executable)
    else:
        # 如果是脚本运行
        return os.path.dirname(os.path.abspath(__file__))