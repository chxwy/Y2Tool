#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
面单名称提示弹窗组件
用于在表格导出完成后显示转换后的面单名称，方便用户复制使用
"""

import tkinter as tk
import re
from datetime import datetime
import os
import sys


class WaybillNameDialog:
    """面单名称提示弹窗"""
    
    def __init__(self, parent, organizer_instance=None):
        self.parent = parent
        self.dialog = None
        self.waybill_names = []
        self.organizer_instance = organizer_instance  # 用于回调通知关闭  # 存储面单名称列表
        # 手动拖动窗口相关状态
        self.is_user_moved = False
        self.user_x = None
        self.user_y = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.start_win_x = 0
        self.start_win_y = 0
        self.user_moved_window = False  # 跟踪用户是否手动移动过窗口
        
        # 抽屉式隐藏功能相关状态
        self.drawer_state = "visible"  # visible, hidden, animating
        self.auto_hide_timer = None
        self.auto_hide_delay = 3000  # 3秒后自动隐藏
        self.hidden_x_offset = None  # 隐藏时的X偏移量
        self.visible_x = None  # 显示时的X位置
        self.animation_steps = 10  # 动画步数
        self.animation_duration = 200  # 动画总时长(ms)
        self.mouse_check_timer = None
        self.edge_detection_width = 50  # 右侧边缘检测宽度
        
        # 配置文件路径
        self.config_file = self._get_config_file_path()
        
    def _get_config_file_path(self):
        """获取配置文件路径"""
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            app_dir = os.path.dirname(sys.executable)
        else:
            # 如果是Python脚本
            app_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(app_dir, "config.json")
    
    def convert_export_name_to_waybill(self, export_name, row_number=None):
        """
        将导出名称转换为面单格式
        例如：急采CHX-10-10-Y2尊祐-床上三件套-艺术家 -> CHX-10.10-3票-Y2面单-尊祐-床上三件套
        
        Args:
            export_name: 导出文件名
            row_number: A列序号（用于票数部分）
        """
        try:
            # 获取配置中心数据
            config = {}
            if self.organizer_instance:
                if hasattr(self.organizer_instance, 'naming_center'):
                    config = self.organizer_instance.naming_center
                elif hasattr(self.organizer_instance, 'config'):
                    config = self.organizer_instance.config.get('naming_center', {})
            
            # 1. 提取缩写（优先从配置获取，否则从文件名识别）
            abbreviation = config.get('business_abbreviation', 'CHX')
            abbreviation_match = re.search(r'([A-Z]{2,4})', export_name)
            if abbreviation_match:
                abbreviation = abbreviation_match.group(1)
            
            # 2. 提取日期部分（如10-10或10.10）
            date_match = re.search(r'(\d{1,2}[-.]?\d{1,2})', export_name)
            if date_match:
                date_part = date_match.group(1).replace('-', '.')
            else:
                # 如果没有找到日期，使用当前日期
                now = datetime.now()
                date_part = f"{now.month}.{now.day}"
            
            # 3. 确定票数：如果提供了行数，使用行数；否则默认为2
            ticket_count = row_number if row_number is not None else 2
            
            # 4. 提取商家名称（如尊祐）
            # 寻找Y2后面的商家名称
            merchant = "尊祐"
            merchant_match = re.search(r'Y2([^-]+)', export_name)
            if merchant_match:
                merchant = merchant_match.group(1)
            else:
                # 尝试从预设列表中匹配
                providers = config.get('logistics_providers', [])
                for p in providers:
                    if p in export_name:
                        # 去掉可能存在的Y1/Y2前缀
                        merchant = p.replace('Y2', '').replace('Y1', '')
                        break
            
            # 5. 提取产品类型和序号
            special_suffixes = config.get('custom_suffixes', ['艺术家', '画家', '设计师'])
            product_type = "窗帘"  # 默认值
            sequence_number = ""  # 序号
            
            # 尝试提取尾部的序号
            sequence_match = re.search(r'-(\d+)$', export_name)
            if sequence_match:
                sequence_number = sequence_match.group(1)
            
            # 匹配：产品类型-特殊后缀-序号 的模式
            for suffix in special_suffixes:
                pattern = rf'-([^-]+)-{suffix}(?:-\d+)?$'
                match = re.search(pattern, export_name)
                if match:
                    product_type = match.group(1)
                    break
            else:
                # 备用逻辑：尝试从分割后的部分中提取
                parts = export_name.split('-')
                if len(parts) >= 3:
                    # 从后往前找第一个非数字、非特殊后缀的部分
                    for i in range(len(parts) - 1, -1, -1):
                        part = parts[i]
                        if not part.isdigit() and part not in special_suffixes and part != 'Y2面单':
                            if not re.match(r'^\d{1,2}$', part) and 'Y2' not in part:
                                product_type = part
                                break
            
            # 构建产品名称部分（产品类型 + 序号）
            product_name_with_sequence = f"{product_type}-{sequence_number}" if sequence_number else product_type
            
            # 6. 使用模板构建面单名称
            template = config.get('waybill_template', '{abbreviation}-{date}-{tickets}票-Y2面单-{merchant}-{product}')
            waybill_name = template.format(
                abbreviation=abbreviation,
                date=date_part,
                tickets=ticket_count,
                merchant=merchant,
                product=product_name_with_sequence
            )
            
            return waybill_name
            
        except Exception as e:
            print(f"转换面单名称时出错: {e}")
            # 如果转换失败，返回一个基础格式
            ticket_count = row_number if row_number is not None else 2
            return f"CHX-{datetime.now().month}.{datetime.now().day}-{ticket_count}票-Y2面单-尊祐-窗帘"
    
    def add_waybill_to_existing(self, export_name, row_number=None):
        """向已存在的弹窗添加新的面单名称"""
        waybill_name = self.convert_export_name_to_waybill(export_name, row_number)
        
        # 如果弹窗不存在，添加到列表并显示
        if not self.dialog or not self.dialog.winfo_exists():
            self.waybill_names.append(waybill_name)
            self.show_waybill_dialog()
            return waybill_name
        
        # 如果弹窗已存在，添加到列表并刷新显示
        self.waybill_names.append(waybill_name)
        self._refresh_dialog_content()
        return waybill_name
    
    def _auto_resize_window(self):
        """自适应调整窗口大小以适应内容"""
        if not self.dialog or not self.dialog.winfo_exists():
            return
        
        # 立即隐藏窗口，避免在调整大小时显示移动效果
        self.dialog.withdraw()
        
        # 强制更新布局，确保所有组件都已正确渲染
        self.dialog.update_idletasks()
        
        # 获取主框架的实际需求尺寸
        main_frame = None
        for widget in self.dialog.winfo_children():
            if isinstance(widget, tk.Frame):
                main_frame = widget
                break
        
        if main_frame:
            # 获取内容的实际需求尺寸
            main_frame.update_idletasks()
            required_width = main_frame.winfo_reqwidth()
            required_height = main_frame.winfo_reqheight()
            
            # 添加一些边距以确保内容完全可见
            margin = 10
            final_width = max(360, required_width + margin)  # 最小宽度360
            final_height = required_height + margin
            
            # 获取当前窗口位置
            current_x = self.dialog.winfo_x()
            current_y = self.dialog.winfo_y()
            
            # 重新计算智能位置（基于新尺寸）
            new_x, new_y = self._calculate_smart_position(final_width, final_height)
            
            # 如果用户没有手动移动过窗口，使用智能位置
            if not hasattr(self, 'user_moved_window') or not self.user_moved_window:
                final_x, final_y = new_x, new_y
            else:
                # 用户移动过窗口，保持当前位置但确保可见
                screen_width = self.dialog.winfo_screenwidth()
                screen_height = self.dialog.winfo_screenheight()
                margin = 3
                final_x = max(margin, min(current_x, screen_width - final_width - margin))
                final_y = max(margin, min(current_y, screen_height - final_height - margin))
            
            # 应用新的窗口大小和位置
            self.dialog.geometry(f"{final_width}x{final_height}+{final_x}+{final_y}")
            
            # 设置好位置后再显示窗口，避免移动效果
            self.dialog.deiconify()

            # 当窗口重新显示时，显式恢复抽屉状态为可见并重置计时器
            self.drawer_state = "visible"
            self._stop_mouse_detection()
            self._start_auto_hide_timer()

    def _refresh_dialog_content(self):
        """刷新弹窗内容以显示新添加的面单名称"""
        if not self.dialog or not self.dialog.winfo_exists():
            return
        
        # 销毁现有内容并重新创建
        for widget in self.dialog.winfo_children():
            widget.destroy()
        
        # 重新创建弹窗内容
        self._create_dialog_content()
        
        # 自适应调整窗口大小
        self._auto_resize_window()
    
    def _create_dialog_content(self):
        """创建弹窗内容（从show_waybill_dialog中提取的公共部分）"""
        # 初始化 Entry 存储列表
        self.waybill_entries = []
        
        # 添加圆角和阴影效果的背景框架
        main_frame = tk.Frame(self.dialog, 
                             bg='#ffffff', 
                             relief='flat',
                             bd=0)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # 添加顶部装饰条
        top_bar = tk.Frame(main_frame, bg='#3498db', height=8)
        top_bar.pack(fill=tk.X)
        # 顶部装饰条作为拖动手柄
        top_bar.bind('<ButtonPress-1>', self._start_drag_with_timer_reset)
        top_bar.bind('<B1-Motion>', self._on_drag)
        top_bar.bind('<ButtonRelease-1>', self._stop_drag)
        
        # 内容框架
        content_frame = tk.Frame(main_frame, bg='#ffffff', padx=15, pady=12)
        content_frame.pack(fill=tk.X, pady=0)
        # 扩大拖动区域：在主框架和内容框架上也绑定拖动事件
        for drag_widget in (main_frame, content_frame):
            drag_widget.bind('<ButtonPress-1>', self._start_drag_with_timer_reset)
            drag_widget.bind('<B1-Motion>', self._on_drag)
            drag_widget.bind('<ButtonRelease-1>', self._stop_drag)
        
        # 标题 - 优化样式
        title_label = tk.Label(content_frame, 
                              text="📋 面单名称 (点击文字可直接编辑)", 
                              font=('Microsoft YaHei', 11, 'bold'),
                              bg='#ffffff',
                              fg='#2c3e50')
        title_label.pack(pady=(0, 10))
        # 标题也支持拖动窗口
        title_label.bind('<ButtonPress-1>', self._start_drag_with_timer_reset)
        title_label.bind('<B1-Motion>', self._on_drag)
        title_label.bind('<ButtonRelease-1>', self._stop_drag)
        
        # 面单名称列表 - 优化布局
        for i, waybill_name in enumerate(self.waybill_names):
            name_frame = tk.Frame(content_frame, bg='#ffffff')
            name_frame.pack(fill=tk.X, pady=3)
            
            # 名称输入框 - 取代原本的 Label，支持手动修改
            name_entry = tk.Entry(name_frame,
                                 font=('Consolas', 10),
                                 bg='#fdfdfd',
                                 fg='#34495e',
                                 relief='flat',
                                 highlightthickness=1,
                                 highlightbackground='#ecf0f1',
                                 highlightcolor='#3498db')
            name_entry.insert(0, waybill_name)
            name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            self.waybill_entries.append(name_entry)
            
            # 按钮容器框架
            button_frame = tk.Frame(name_frame, bg='#ffffff')
            button_frame.pack(side=tk.RIGHT)
            
            # - 按钮 - 小巧低调样式 (改为递减逻辑)
            minus_button = tk.Button(button_frame, 
                                       text="➖", 
                                       font=('Segoe UI Emoji', 8),
                                       bg='#f8f9fa',
                                       fg='#6c757d',
                                       relief='flat',
                                       bd=0,
                                       padx=6,
                                       pady=3,
                                       cursor='hand2',
                                       activebackground='#e9ecef',
                                       activeforeground='#495057',
                                       highlightthickness=0,
                                       command=lambda idx=i: self.on_minus_click(idx))
            minus_button.pack(side=tk.LEFT, padx=(0, 4))
            
            # 按钮悬停效果 - 低调的反馈
            def on_minus_enter(e, btn=minus_button):
                btn.config(bg='#e9ecef', fg='#495057')
            def on_minus_leave(e, btn=minus_button):
                btn.config(bg='#f8f9fa', fg='#6c757d')
            
            minus_button.bind('<Enter>', on_minus_enter)
            minus_button.bind('<Leave>', on_minus_leave)
            
            # 复制按钮 - 恢复原始样式
            copy_button = tk.Button(button_frame, 
                                   text="复制", 
                                   font=('Microsoft YaHei', 9),
                                   bg='#3498db',
                                   fg='white',
                                   relief='flat',
                                   bd=0,
                                   padx=12,
                                   pady=4,
                                   cursor='hand2',
                                   command=lambda idx=i: self.on_copy_click(idx))
            copy_button.pack(side=tk.LEFT)
            
            # 按钮悬停效果
            def on_enter(e, btn=copy_button):
                btn.config(bg='#2980b9')
            def on_leave(e, btn=copy_button):
                btn.config(bg='#3498db')
            
            copy_button.bind('<Enter>', on_enter)
            copy_button.bind('<Leave>', on_leave)
        
        # 底部关闭按钮 - 贴底显示，更紧凑的高度
        close_button = tk.Button(content_frame, 
                                text="关闭", 
                                font=('Microsoft YaHei', 9),
                                bg='#95a5a6',
                                fg='white',
                                relief='flat',
                                bd=0,
                                padx=20,
                                pady=2,
                                cursor='hand2',
                                command=self._close_dialog_with_timer_reset)
        close_button.pack(side=tk.BOTTOM, pady=(10, 0))
        
        # 关闭按钮悬停效果
        def on_close_enter(e):
            close_button.config(bg='#7f8c8d')
        def on_close_leave(e):
            close_button.config(bg='#95a5a6')
        
        close_button.bind('<Enter>', on_close_enter)
        close_button.bind('<Leave>', on_close_leave)
        
        # 绑定ESC键关闭
        self.dialog.bind('<Escape>', lambda e: self._close_dialog_with_timer_reset())
        
        # 绑定点击外部关闭（可选）
        self.dialog.bind('<Button-1>', self._on_click_outside)

    def _start_drag(self, event):
        """开始拖动"""
        self.start_x = event.x_root
        self.start_y = event.y_root
        self.start_dialog_x = self.dialog.winfo_x()
        self.start_dialog_y = self.dialog.winfo_y()
    
    def _start_drag_with_timer_reset(self, event):
        """开始拖动并重置自动隐藏计时器"""
        self._reset_auto_hide_timer()
        self._start_drag(event)
    
    def _close_dialog_with_timer_reset(self):
        """关闭对话框并重置自动隐藏计时器"""
        self._reset_auto_hide_timer()
        self.close_dialog()

    def _on_drag(self, event):
        """拖动中"""
        if hasattr(self, 'start_x') and hasattr(self, 'start_y'):
            # 计算鼠标移动的距离
            dx = event.x_root - self.start_x
            dy = event.y_root - self.start_y
            
            # 计算新的窗口位置
            new_x = self.start_dialog_x + dx
            new_y = self.start_dialog_y + dy
            
            # 更新窗口位置
            self.dialog.geometry(f"+{new_x}+{new_y}")
            
            # 标记用户已手动移动窗口
            self.user_moved_window = True

    def _stop_drag(self, event):
        """停止拖动"""
        pass
    
    def _calculate_smart_position(self, window_width, window_height):
        """智能计算窗口位置，确保窗口完全可见且向上延伸"""
        # 获取屏幕尺寸
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        
        # 设置边距
        margin = 3
        taskbar_height = 80  # 任务栏高度估计
        
        # 计算右下角的基础位置
        base_x = screen_width - window_width - margin
        base_y = screen_height - taskbar_height - margin
        
        # 如果窗口高度超出屏幕，向上调整
        if base_y < 0:
            # 窗口太高，调整到屏幕顶部
            final_y = margin
        else:
            # 窗口从底部向上延伸
            final_y = base_y - window_height
            
            # 确保窗口不会超出屏幕顶部
            if final_y < margin:
                final_y = margin
        
        # 确保窗口不会超出屏幕右侧
        final_x = min(base_x, screen_width - window_width - margin)
        
        return final_x, final_y

    def add_waybill_name(self, export_name, row_number=None):
        """添加一个面单名称到列表"""
        waybill_name = self.convert_export_name_to_waybill(export_name, row_number)
        self.waybill_names.append(waybill_name)
        return waybill_name
    
    def show_waybill_dialog(self):
        """显示面单名称弹窗"""
        if not self.waybill_names:
            return
        
        # 如果弹窗已存在，刷新内容
        if self.dialog and self.dialog.winfo_exists():
            self._refresh_dialog_content()
            return
        
        # 创建弹窗
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("面单名称提示")
        
        # 设置窗口属性
        self.dialog.attributes('-topmost', True)  # 始终置顶
        self.dialog.attributes('-alpha', 0.88)    # 透明度88%
        self.dialog.resizable(False, False)
        self.dialog.overrideredirect(True)        # 去除标题栏，更简洁
        
        # 初始窗口大小 - 先设置一个临时大小
        initial_width = 360
        initial_height = 200  # 临时高度，后面会自适应调整
        
        # 使用智能位置计算（基于临时尺寸）
        x, y = self._calculate_smart_position(initial_width, initial_height)
        
        self.dialog.geometry(f"{initial_width}x{initial_height}+{x}+{y}")
        
        # 创建弹窗内容
        self._create_dialog_content()

        # 绑定鼠标进入/移动/离开事件以控制计时器
        self.dialog.bind('<Enter>', lambda e: self._reset_auto_hide_timer())
        self.dialog.bind('<Motion>', lambda e: self._reset_auto_hide_timer())
        self.dialog.bind('<Leave>', lambda e: self._immediate_hide_if_at_edge())
        
        # 自适应窗口大小 - 基于实际内容测量
        self._auto_resize_window()
        
        # 绑定ESC键关闭
        self.dialog.bind('<Escape>', lambda e: self.close_dialog())
        
        # 设置焦点
        self.dialog.focus_force()
        
        # 取消自动关闭，改为需要用户手动关闭
        # self.dialog.after(30000, self.close_dialog)
    
    def _on_click_outside(self, event):
        """点击弹窗外部时关闭（可选功能）"""
        # 这里可以添加点击外部关闭的逻辑
        pass
    
    def copy_only(self, waybill_name):
        """仅复制面单名称到剪贴板"""
        try:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(waybill_name)
            self.parent.update()  # 确保剪贴板更新
        except Exception as e:
            pass
    
    def copy_and_close(self, waybill_name):
        """复制面单名称并关闭弹窗"""
        try:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(waybill_name)
            self.parent.update()  # 确保剪贴板更新
            
            # 关闭弹窗
            self.close_dialog()
        except Exception as e:
            pass

    def copy_and_remove(self, waybill_name):
        """复制面单名称并从列表中移除"""
        try:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(waybill_name)
            self.parent.update()  # 确保剪贴板更新
            
            # 从列表中移除
            if waybill_name in self.waybill_names:
                self.waybill_names.remove(waybill_name)
                
                # 如果列表为空，关闭弹窗
                if not self.waybill_names:
                    self.close_dialog()
                else:
                    # 否则刷新弹窗内容
                    self._refresh_dialog_content()
        except Exception as e:
            pass

    def on_copy_click(self, index):
        """复制按钮点击：抓取当前输入框内容，多行逐行消失，单行关闭弹窗"""
        # 重置自动隐藏计时器
        self._reset_auto_hide_timer()
        
        # 抓取当前 Entry 中的内容
        if 0 <= index < len(self.waybill_entries):
            current_name = self.waybill_entries[index].get().strip()
            
            # 单行则复制并关闭；多行则复制并移除该行
            if len(self.waybill_names) <= 1:
                self.copy_and_close(current_name)
            else:
                self.copy_and_remove(current_name)
    
    def on_minus_click(self, index):
        """处理减号按钮点击事件：抓取当前输入框内容并递减或移除序号"""
        # 重置自动隐藏计时器
        self._reset_auto_hide_timer()
        
        if 0 <= index < len(self.waybill_entries):
            # 获取当前输入框中的实时内容
            waybill_name = self.waybill_entries[index].get().strip()
            
            # 实现名称递减逻辑
            new_name = self._decrement_waybill_name(waybill_name)
            
            # 更新列表中的名称
            self.waybill_names[index] = new_name
            
            # 刷新弹窗内容以显示更新后的名称
            self._refresh_dialog_content()
    
    def _decrement_waybill_name(self, name):
        """递减面单名称，如果末尾是-2则移除变为原名，如果是-3及以上则递减数字"""
        import re
        
        # 检查名称末尾是否已有数字（格式：-数字）
        match = re.search(r'-(\d+)$', name)
        
        if match:
            current_number = int(match.group(1))
            if current_number > 2:
                # 如果数字大于2，递减
                new_number = current_number - 1
                new_name = re.sub(r'-\d+$', f'-{new_number}', name)
            else:
                # 如果数字是2，移除-2，恢复原名
                new_name = re.sub(r'-2$', '', name)
        else:
            # 如果末尾没有数字，保持不变（或者根据需求也可以不处理）
            new_name = name
        
        return new_name
    
    def close_dialog(self):
        """关闭弹窗"""
        # 清理抽屉式隐藏相关的计时器
        self._cancel_auto_hide_timer()
        self._stop_mouse_detection()
        
        if self.dialog:
            self.dialog.destroy()
            self.dialog = None
        
        # 重置抽屉状态
        self.drawer_state = "visible"
        
        # 清空面单名称列表
        self.waybill_names = []
    
    def show_single_waybill(self, export_name, row_number=None):
        """显示单个面单名称（便捷方法）"""
        waybill_name = self.convert_export_name_to_waybill(export_name, row_number)
        self.waybill_names = [waybill_name]
        self.show_waybill_dialog()
    
    def show_multiple_waybills(self, waybill_names):
        """显示多个面单名称（便捷方法）"""
        # 如果弹窗已存在，追加新的面单名称而不是覆盖
        if self.dialog and self.dialog.winfo_exists():
            # 追加新的面单名称到现有列表
            self.waybill_names.extend(waybill_names)
            # 刷新弹窗内容
            self._refresh_dialog_content()
        else:
            # 如果弹窗不存在，直接设置面单名称列表
            self.waybill_names = waybill_names
            self.show_waybill_dialog()
    
    # ==================== 抽屉式隐藏功能 ====================

    def _is_window_at_right_edge(self):
        """检查窗口是否贴着屏幕右侧边缘"""
        if not self.dialog or not self.dialog.winfo_exists():
            return False
        try:
            window_x = self.dialog.winfo_x()
            window_width = self.dialog.winfo_width()
            screen_width = self.dialog.winfo_screenwidth()
            
            # 计算窗口右边缘位置
            window_right_edge = window_x + window_width
            
            # 允许一定的误差范围（比如10像素），认为是贴着右边缘
            edge_tolerance = 10
            
            return abs(window_right_edge - screen_width) <= edge_tolerance
        except:
            return False

    def _is_mouse_inside_window(self):
        """检查鼠标是否位于当前弹窗内部"""
        if not self.dialog or not self.dialog.winfo_exists():
            return False
        try:
            mouse_x = self.dialog.winfo_pointerx()
            mouse_y = self.dialog.winfo_pointery()
            win_x = self.dialog.winfo_rootx()
            win_y = self.dialog.winfo_rooty()
            win_w = self.dialog.winfo_width()
            win_h = self.dialog.winfo_height()
            return win_x <= mouse_x <= win_x + win_w and win_y <= mouse_y <= win_y + win_h
        except:
            return False

    
    def _immediate_hide_if_at_edge(self):
        """鼠标离开时立即隐藏（仅当窗口贴着右侧边缘时）"""
        if self.drawer_state == "visible" and self._is_window_at_right_edge():
            mouse_inside = self._is_mouse_inside_window()
            # 计算鼠标是否仍处于窗口右侧的检测区域内（即可以再次唤醒窗口的热区），
            # 如果鼠标仍在该热区，则不立即隐藏，而是重置自动隐藏计时器。
            try:
                mouse_x = self.dialog.winfo_pointerx()
                mouse_y = self.dialog.winfo_pointery()
                screen_width = self.dialog.winfo_screenwidth()
                window_y = self.dialog.winfo_y()
                window_height = self.dialog.winfo_height()

                detection_left = screen_width - self.edge_detection_width  # 与 _check_mouse_position 保持一致
                detection_top = window_y - 20
                detection_bottom = window_y + window_height + 20

                mouse_near_edge = (mouse_x >= detection_left and
                                   detection_top <= mouse_y <= detection_bottom)
            except Exception:
                mouse_near_edge = False

            if not mouse_inside and not mouse_near_edge:
                # 鼠标既不在窗口内部，也不在右侧检测热区，执行隐藏
                self._hide_to_drawer()
            else:
                # 鼠标仍在窗口或热区，重置计时器，防止出现反复隐藏/显示的抖动
                self._reset_auto_hide_timer()

    def _start_auto_hide_timer(self):
        """启动自动隐藏计时器"""
        self._cancel_auto_hide_timer()
        if self.drawer_state == "visible":
            self.auto_hide_timer = self.dialog.after(self.auto_hide_delay, self._auto_hide_to_drawer)
    
    def _cancel_auto_hide_timer(self):
        """取消自动隐藏计时器"""
        if self.auto_hide_timer:
            self.dialog.after_cancel(self.auto_hide_timer)
            self.auto_hide_timer = None
    
    def _reset_auto_hide_timer(self):
        """重置自动隐藏计时器（用户活动时调用）"""
        if self.drawer_state == "visible":
            self._start_auto_hide_timer()
    
    def _auto_hide_to_drawer(self):
        """自动隐藏到右侧抽屉"""
        if self.drawer_state == "visible" and self.dialog and self.dialog.winfo_exists():
            # 只有当窗口贴着右侧边缘时才进行自动隐藏
            if not self._is_window_at_right_edge():
                return
                
            if self._is_mouse_inside_window():
                # 鼠标仍在窗口内部，重新计时而不隐藏
                self._start_auto_hide_timer()
            else:
                self._hide_to_drawer()
    
    def _hide_to_drawer(self):
        """隐藏窗口到右侧抽屉"""
        if self.drawer_state != "visible" or not self.dialog or not self.dialog.winfo_exists():
            return
        
        self.drawer_state = "animating"
        self._cancel_auto_hide_timer()
        
        # 记录当前可见位置
        self.visible_x = self.dialog.winfo_x()
        current_y = self.dialog.winfo_y()
        window_width = self.dialog.winfo_width()
        screen_width = self.dialog.winfo_screenwidth()
        
        # 计算隐藏位置（只露出一小部分）
        visible_edge_width = 20  # 露出的边缘宽度
        self.hidden_x_offset = screen_width - visible_edge_width
        
        # 执行滑动动画
        self._animate_to_position(self.hidden_x_offset, current_y, self._on_hide_complete)
    
    def _show_from_drawer(self):
        """从右侧抽屉显示窗口"""
        if self.drawer_state != "hidden" or not self.dialog or not self.dialog.winfo_exists():
            return
        
        self.drawer_state = "animating"
        current_y = self.dialog.winfo_y()
        
        # 恢复到可见位置
        target_x = self.visible_x if self.visible_x is not None else self.dialog.winfo_screenwidth() - self.dialog.winfo_width() - 50
        
        # 执行滑动动画
        self._animate_to_position(target_x, current_y, self._on_show_complete)
    
    def _animate_to_position(self, target_x, target_y, callback=None):
        """平滑动画到目标位置"""
        if not self.dialog or not self.dialog.winfo_exists():
            return
        
        start_x = self.dialog.winfo_x()
        start_y = self.dialog.winfo_y()
        
        step_x = (target_x - start_x) / self.animation_steps
        step_y = (target_y - start_y) / self.animation_steps
        step_delay = self.animation_duration // self.animation_steps
        
        def animate_step(step):
            if not self.dialog or not self.dialog.winfo_exists():
                return
            
            if step < self.animation_steps:
                new_x = int(start_x + step_x * step)
                new_y = int(start_y + step_y * step)
                self.dialog.geometry(f"+{new_x}+{new_y}")
                self.dialog.after(step_delay, lambda: animate_step(step + 1))
            else:
                # 动画完成，设置最终位置
                self.dialog.geometry(f"+{int(target_x)}+{int(target_y)}")
                if callback:
                    callback()
        
        animate_step(0)
    
    def _on_hide_complete(self):
        """隐藏动画完成回调"""
        self.drawer_state = "hidden"
        self._start_mouse_detection()
    
    def _on_show_complete(self):
        """显示动画完成回调"""
        self.dialog.deiconify()
    
        # 当窗口重新显示（例如刷新内容或新面单添加）时，
        # 若之前处于隐藏状态，需要显式将抽屉状态恢复为可见，
        # 并停止隐藏状态下的鼠标检测逻辑。
        self.drawer_state = "visible"
        self._stop_mouse_detection()
        
        self._start_auto_hide_timer()
    
    def _start_mouse_detection(self):
        """启动鼠标检测（用于从隐藏状态唤醒）"""
        if self.drawer_state == "hidden":
            self._check_mouse_position()
    
    def _stop_mouse_detection(self):
        """停止鼠标检测"""
        if self.mouse_check_timer:
            self.dialog.after_cancel(self.mouse_check_timer)
            self.mouse_check_timer = None
    
    def _check_mouse_position(self):
        """检查鼠标位置，判断是否需要唤醒窗口"""
        if self.drawer_state != "hidden" or not self.dialog or not self.dialog.winfo_exists():
            return
        
        try:
            # 获取鼠标位置
            mouse_x = self.dialog.winfo_pointerx()
            mouse_y = self.dialog.winfo_pointery()
            screen_width = self.dialog.winfo_screenwidth()
            
            # 获取窗口隐藏时的位置信息
            hidden_x = screen_width + self.hidden_x_offset  # 窗口隐藏时的X位置
            window_y = self.dialog.winfo_y()  # 窗口的Y位置
            window_height = self.dialog.winfo_height()  # 窗口高度
            
            # 检查鼠标是否在窗口右侧的检测区域内
            # X坐标：屏幕右边缘向左50像素的区域
            # Y坐标：窗口的垂直范围内（上下各扩展20像素）
            detection_left = screen_width - self.edge_detection_width
            detection_top = window_y - 20
            detection_bottom = window_y + window_height + 20
            
            if (mouse_x >= detection_left and 
                mouse_y >= detection_top and 
                mouse_y <= detection_bottom):
                self._show_from_drawer()
                return
            
            # 继续检测
            self.mouse_check_timer = self.dialog.after(100, self._check_mouse_position)
        except:
            # 如果出错，继续检测
            self.mouse_check_timer = self.dialog.after(100, self._check_mouse_position)