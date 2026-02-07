# -*- coding: utf-8 -*-
"""
远程更新模块 - Y2订单处理辅助工具
功能：检查更新、下载更新包、启动更新助手
"""

import os
import sys
import json
import hashlib
import tempfile
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from urllib.parse import urlparse
import time

# 版本信息
CURRENT_VERSION = "1.9.0"
# GitHub 主更新源（使用 GitHub Pages 或 raw 方式）
VERSION_CHECK_URL = "https://raw.githubusercontent.com/chxwy/Y2Tool/main/docs/version.json"
# 备用更新源（可以换成 Gitee 或其他镜像）
BACKUP_CHECK_URL = "https://gitee.com/chxwy/Y2Tool/raw/main/docs/version.json"


class UpdateChecker:
    """更新检查器"""
    
    def __init__(self):
        self.latest_version = None
        self.download_url = None
        self.changelog = []
        self.force_update = False
        self.file_size = 0
        self.file_hash = None
        self.error_msg = None
        
    def check_update(self, use_backup=False):
        """
        检查是否有新版本
        返回: (has_update: bool, version_info: dict)
        """
        try:
            import requests
            
            url = BACKUP_CHECK_URL if use_backup else VERSION_CHECK_URL
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            version_info = response.json()
            self.latest_version = version_info.get('version', '0.0.0')
            self.download_url = version_info.get('download_url', '')
            self.changelog = version_info.get('changelog', [])
            self.force_update = version_info.get('force_update', False)
            self.file_size = version_info.get('file_size', 0)
            self.file_hash = version_info.get('hash', '')
            
            # 版本号比较
            has_update = self._compare_version(CURRENT_VERSION, self.latest_version)
            
            return has_update, version_info
            
        except Exception as e:
            self.error_msg = str(e)
            # 如果主源失败，尝试备用源
            if not use_backup:
                return self.check_update(use_backup=True)
            return False, None
    
    def _compare_version(self, current, latest):
        """比较版本号，返回 True 如果有新版本"""
        try:
            current_parts = [int(x) for x in current.split('.')]
            latest_parts = [int(x) for x in latest.split('.')]
            
            # 补齐版本号位数
            while len(current_parts) < len(latest_parts):
                current_parts.append(0)
            while len(latest_parts) < len(current_parts):
                latest_parts.append(0)
            
            for i in range(len(current_parts)):
                if latest_parts[i] > current_parts[i]:
                    return True
                elif latest_parts[i] < current_parts[i]:
                    return False
            return False
        except:
            return False
    
    def download_update(self, download_path, progress_callback=None):
        """
        下载更新包
        progress_callback: 回调函数(current_size, total_size)
        """
        try:
            import requests
            
            response = requests.get(self.download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            if total_size == 0:
                total_size = self.file_size
            
            downloaded = 0
            with open(download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback:
                            progress_callback(downloaded, total_size)
            
            # 验证文件哈希
            if self.file_hash:
                file_hash = self._calculate_hash(download_path)
                if not file_hash.startswith(self.file_hash.split(':')[-1][:16]):
                    os.remove(download_path)
                    return False, "文件校验失败"
            
            return True, None
            
        except Exception as e:
            if os.path.exists(download_path):
                os.remove(download_path)
            return False, str(e)
    
    def _calculate_hash(self, file_path):
        """计算文件 SHA256 哈希"""
        sha256 = hashlib.sha256()
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()


class UpdateDialog:
    """更新提示对话框"""
    
    def __init__(self, parent, version_info, checker):
        self.checker = checker
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"发现新版本 - {version_info['version']}")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 500) // 2
        y = (self.dialog.winfo_screenheight() - 400) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self._create_ui(version_info)
        
    def _create_ui(self, version_info):
        """创建对话框UI"""
        # 标题
        title_frame = ttk.Frame(self.dialog, padding="20")
        title_frame.pack(fill='x')
        
        ttk.Label(
            title_frame,
            text="🎉 发现新版本",
            font=('Microsoft YaHei UI', 16, 'bold'),
            foreground='#2E86AB'
        ).pack()
        
        ttk.Label(
            title_frame,
            text=f"当前版本: {CURRENT_VERSION}  →  最新版本: {version_info['version']}",
            font=('Microsoft YaHei UI', 10)
        ).pack(pady=(10, 0))
        
        # 更新日志
        log_frame = ttk.LabelFrame(self.dialog, text="更新内容", padding="10")
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        log_text = tk.Text(
            log_frame,
            wrap='word',
            font=('Microsoft YaHei UI', 10),
            height=10,
            padx=5,
            pady=5
        )
        log_text.pack(fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=log_text.yview)
        scrollbar.pack(side='right', fill='y')
        log_text.configure(yscrollcommand=scrollbar.set)
        
        # 填充更新日志
        changelog = version_info.get('changelog', [])
        if changelog:
            for item in changelog:
                log_text.insert('end', f"• {item}\n")
        else:
            log_text.insert('end', "暂无更新说明")
        log_text.configure(state='disabled')
        
        # 进度条（初始隐藏）
        self.progress_frame = ttk.Frame(self.dialog)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(pady=5)
        self.progress_label = ttk.Label(self.progress_frame, text="准备下载...")
        self.progress_label.pack()
        
        # 按钮
        self.button_frame = ttk.Frame(self.dialog, padding="20")
        self.button_frame.pack(fill='x')
        
        self.update_btn = ttk.Button(
            self.button_frame,
            text="立即更新",
            command=self._start_update
        )
        self.update_btn.pack(side='left', padx=(0, 10))
        
        self.later_btn = ttk.Button(
            self.button_frame,
            text="稍后提醒",
            command=self._remind_later
        )
        self.later_btn.pack(side='left', padx=(0, 10))
        
        if not version_info.get('force_update', False):
            self.skip_btn = ttk.Button(
                self.button_frame,
                text="跳过此版本",
                command=self._skip_version
            )
            self.skip_btn.pack(side='right')
        
    def _start_update(self):
        """开始更新"""
        self.update_btn.configure(state='disabled')
        self.later_btn.configure(state='disabled')
        if hasattr(self, 'skip_btn'):
            self.skip_btn.configure(state='disabled')
        
        self.button_frame.pack_forget()
        self.progress_frame.pack(fill='x', padx=20, pady=10)
        
        # 在后台线程下载
        threading.Thread(target=self._download_and_install, daemon=True).start()
    
    def _download_and_install(self):
        """下载并安装更新"""
        try:
            # 创建临时目录
            temp_dir = tempfile.gettempdir()
            download_path = os.path.join(temp_dir, f"Y2订单处理辅助工具_update_{self.checker.latest_version}.zip")
            
            # 下载更新包
            def progress_callback(current, total):
                if total > 0:
                    percent = (current / total) * 100
                    self.progress_var.set(percent)
                    self.progress_label.configure(
                        text=f"下载中... {current//1024//1024}MB / {total//1024//1024}MB ({percent:.1f}%)"
                    )
                self.dialog.update_idletasks()
            
            success, error = self.checker.download_update(download_path, progress_callback)
            
            if not success:
                self.dialog.after(0, lambda: self._show_error(f"下载失败: {error}"))
                return
            
            self.progress_label.configure(text="下载完成，准备安装...")
            
            # 启动更新助手
            self._launch_updater(download_path)
            
            self.result = 'update'
            self.dialog.after(0, self.dialog.destroy)
            
        except Exception as e:
            self.dialog.after(0, lambda: self._show_error(str(e)))
    
    def _launch_updater(self, update_package_path):
        """启动更新助手程序"""
        try:
            # 获取当前程序路径
            if getattr(sys, 'frozen', False):
                # PyInstaller 打包后的路径
                current_dir = os.path.dirname(sys.executable)
                # 如果是 onefile 模式，sys.executable 就是主程序
                # 如果是 onedir 模式，sys.executable 在 _internal 或同级目录
                if '_internal' in current_dir:
                    current_dir = os.path.dirname(current_dir)
            else:
                # 开发环境
                current_dir = os.path.dirname(os.path.abspath(__file__))
            
            # 更新助手路径
            updater_path = os.path.join(current_dir, 'updater.exe')
            
            # 如果更新助手不存在，使用内置方法
            if not os.path.exists(updater_path):
                updater_path = os.path.join(current_dir, '_internal', 'updater.exe')
            
            # 启动更新助手
            if os.path.exists(updater_path):
                subprocess.Popen([
                    updater_path,
                    update_package_path,
                    current_dir,
                    sys.executable if getattr(sys, 'frozen', False) else ''
                ], shell=False)
            else:
                # 如果没有独立的更新助手，使用 Python 脚本方式
                updater_script = os.path.join(current_dir, 'updater.py')
                if os.path.exists(updater_script):
                    subprocess.Popen([
                        sys.executable,
                        updater_script,
                        update_package_path,
                        current_dir,
                        sys.executable if getattr(sys, 'frozen', False) else ''
                    ], shell=False)
                else:
                    # 最后手段：直接解压并提示用户手动重启
                    self._extract_and_notify(update_package_path, current_dir)
                    
        except Exception as e:
            print(f"启动更新助手失败: {e}")
    
    def _extract_and_notify(self, zip_path, target_dir):
        """解压并通知用户手动重启"""
        import zipfile
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(target_dir)
            messagebox.showinfo(
                "更新完成",
                "更新文件已下载并解压完成。\n请手动重启程序以应用更新。",
                parent=self.dialog
            )
        except Exception as e:
            messagebox.showerror(
                "更新失败",
                f"解压更新文件失败: {e}\n请手动下载更新。",
                parent=self.dialog
            )
    
    def _show_error(self, message):
        """显示错误信息"""
        messagebox.showerror("更新失败", message, parent=self.dialog)
        self.result = 'error'
        self.dialog.destroy()
    
    def _remind_later(self):
        """稍后提醒"""
        self.result = 'later'
        self.dialog.destroy()
    
    def _skip_version(self):
        """跳过此版本"""
        # 保存跳过的版本号到配置文件
        self._save_skip_version(self.checker.latest_version)
        self.result = 'skip'
        self.dialog.destroy()
    
    def _save_skip_version(self, version):
        """保存跳过的版本号"""
        try:
            config_path = os.path.join(
                os.path.expanduser('~'),
                '.Y2订单处理辅助工具',
                'update_config.json'
            )
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            
            config = {}
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            config['skipped_version'] = version
            config['skip_time'] = time.time()
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def show(self):
        """显示对话框并等待结果"""
        self.dialog.wait_window()
        return self.result


def check_for_updates(parent=None, silent=False):
    """
    检查更新的入口函数
    
    Args:
        parent: 父窗口
        silent: 是否静默检查（无更新时不提示）
    
    Returns:
        bool: True 如果有更新且用户选择更新
    """
    checker = UpdateChecker()
    has_update, version_info = checker.check_update()
    
    if not has_update:
        if not silent:
            messagebox.showinfo("检查更新", "当前已是最新版本！", parent=parent)
        return False
    
    if version_info is None:
        if not silent:
            messagebox.showwarning(
                "检查更新",
                f"检查更新失败: {checker.error_msg}\n请检查网络连接。",
                parent=parent
            )
        return False
    
    # 检查是否跳过了此版本
    if _is_version_skipped(version_info['version']):
        return False
    
    # 显示更新对话框
    dialog = UpdateDialog(parent, version_info, checker)
    result = dialog.show()
    
    return result == 'update'


def _is_version_skipped(version):
    """检查用户是否跳过了此版本"""
    try:
        config_path = os.path.join(
            os.path.expanduser('~'),
            '.Y2订单处理辅助工具',
            'update_config.json'
        )
        
        if not os.path.exists(config_path):
            return False
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        skipped = config.get('skipped_version')
        skip_time = config.get('skip_time', 0)
        
        # 7天内跳过的版本不再提示
        if skipped == version and (time.time() - skip_time) < 7 * 24 * 3600:
            return True
        
        return False
    except:
        return False


if __name__ == '__main__':
    # 测试
    root = tk.Tk()
    root.withdraw()
    check_for_updates(root)
