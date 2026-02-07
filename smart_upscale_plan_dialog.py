import tkinter as tk
from tkinter import ttk, messagebox
import os

# --- 与主程序一致的UI主题常量 ---
# 使用与主程序相同的配色方案
BG_COLOR = "#FFFFFF"  # 主背景色
CARD_BG = "#FFFFFF"   # 卡片背景色
SIDEBAR_BG = "#F8F9FA"  # 侧边栏背景色
TEXT_COLOR = "#212529"  # 主文本色
TEXT_SECONDARY = "#6C757D"  # 次要文本色
PRIMARY_COLOR = "#0D6EFD"  # 主色调
PRIMARY_HOVER = "#0B5ED7"  # 主色调悬停
SUCCESS_COLOR = "#198754"  # 成功色
SUCCESS_HOVER = "#157347"  # 成功色悬停
CANCEL_COLOR = "#6C757D"   # 取消按钮色
CANCEL_HOVER = "#5C636A"   # 取消按钮悬停色
ACCENT_COLOR = "#FD7E14"   # 强调色
SHADOW_COLOR = "#DEE2E6"   # 阴影色
BORDER_COLOR = "#DEE2E6"   # 边框色
FONT_FAMILY = "Microsoft YaHei UI"

def create_gradient_frame(parent, width, height, color1, color2):
    """创建渐变背景的Canvas"""
    canvas = tk.Canvas(parent, width=width, height=height, highlightthickness=0)
    
    # 创建渐变效果
    steps = 100
    for i in range(steps):
        # 计算渐变颜色
        ratio = i / steps
        r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
        r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
        
        r = int(r1 + (r2 - r1) * ratio)
        g = int(g1 + (g2 - g1) * ratio)
        b = int(b1 + (b2 - b1) * ratio)
        
        color = f"#{r:02x}{g:02x}{b:02x}"
        y = int(height * ratio)
        canvas.create_rectangle(0, y, width, y + height // steps + 1, fill=color, outline=color)
    
    return canvas

def create_card_frame(parent, **kwargs):
    """创建带阴影效果的卡片框架"""
    # 外层阴影框架
    shadow_frame = tk.Frame(parent, bg=SHADOW_COLOR, **kwargs)
    
    # 内层卡片框架
    card_frame = tk.Frame(shadow_frame, bg=CARD_BG, padx=20, pady=15)
    card_frame.pack(padx=3, pady=3, fill=tk.BOTH, expand=True)
    
    return shadow_frame, card_frame

def show_smart_upscale_plan_dialog(parent, plan_data, start_callback=None):
    """显示现代化智能倍数匹配弹窗"""
    dialog = tk.Toplevel(parent)
    dialog.title("🎯 智能倍数匹配")
    
    # 固定窗口大小，不随内容变化
    screen_height = dialog.winfo_screenheight()
    max_height = int(screen_height * 0.8)
    dialog_height = min(580, max_height)
    dialog_width = 800
    
    dialog.geometry(f"{dialog_width}x{dialog_height}")
    dialog.resizable(False, False)
    dialog.minsize(dialog_width, dialog_height)
    dialog.maxsize(dialog_width, dialog_height)
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    dialog.focus_force()
    dialog.configure(bg=BG_COLOR)

    # 存储修改后的数据
    modified_plan_data = plan_data.copy()
    modified_plan_data['processing_list'] = [item.copy() for item in plan_data.get('processing_list', [])]

    # --- 现代化样式配置 ---
    style = ttk.Style(dialog)
    style.theme_use("clam")

    # 标签样式 - 与主程序保持一致
    style.configure("Modern.TLabel", 
                    background=CARD_BG, 
                    foreground=TEXT_COLOR, 
                    font=(FONT_FAMILY, 11))
    style.configure("Title.TLabel", 
                    font=(FONT_FAMILY, 16, "bold"),
                    foreground=TEXT_COLOR,  # 使用主文本色
                    background=BG_COLOR)
    style.configure("Subtitle.TLabel", 
                    font=(FONT_FAMILY, 10),
                    foreground=TEXT_SECONDARY,
                    background=BG_COLOR)
    style.configure("Stats.TLabel", 
                    font=(FONT_FAMILY, 10),
                    foreground=TEXT_COLOR,
                    background=BG_COLOR)
    style.configure("StatsValue.TLabel", 
                    font=(FONT_FAMILY, 14, "bold"),
                    foreground=PRIMARY_COLOR,
                    background=BG_COLOR)
    
    # 现代化表格样式 - 提高清晰度
    style.configure("Modern.Treeview", 
                    background=CARD_BG, 
                    foreground=TEXT_COLOR, 
                    fieldbackground=CARD_BG,
                    rowheight=34,  # 稍微增加行高提高清晰度
                    font=(FONT_FAMILY, 10),
                    borderwidth=1,
                    relief="solid")
    style.map("Modern.Treeview", 
              background=[('selected', PRIMARY_COLOR)],
              foreground=[('selected', 'white')])

    style.configure("Modern.Treeview.Heading", 
                    font=(FONT_FAMILY, 11, "bold"), 
                    background=PRIMARY_COLOR, 
                    foreground="white",
                    relief="flat",
                    borderwidth=1)
    style.map("Modern.Treeview.Heading", 
              background=[('active', PRIMARY_HOVER)])

    # 优化滚动条样式 - 更清晰
    style.configure("Modern.Vertical.TScrollbar",
                    background=SHADOW_COLOR,
                    troughcolor=SIDEBAR_BG,
                    borderwidth=1,
                    arrowcolor=TEXT_SECONDARY,
                    darkcolor=BORDER_COLOR,
                    lightcolor=SIDEBAR_BG,
                    relief="solid")
    style.map("Modern.Vertical.TScrollbar",
              background=[('active', TEXT_SECONDARY), ('pressed', TEXT_COLOR)])

    # 主容器
    main_container = tk.Frame(dialog, bg=BG_COLOR)
    main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    # === 清晰的标题区域 ===
    title_container = tk.Frame(main_container, bg=BG_COLOR)
    title_container.pack(fill=tk.X, pady=(0, 15))

    # 主标题 - 更清晰
    title_label = tk.Label(title_container, text="🎯 智能倍数匹配", 
                          font=(FONT_FAMILY, 16, "bold"), 
                          bg=BG_COLOR, fg=TEXT_COLOR)
    title_label.pack(side=tk.LEFT)

    # 副标题
    subtitle_label = tk.Label(title_container, text="智能分析图片尺寸，自动匹配最佳放大倍数", 
                             font=(FONT_FAMILY, 10), 
                             bg=BG_COLOR, fg=TEXT_SECONDARY)
    subtitle_label.pack(side=tk.LEFT, padx=(10, 0))

    # === 统计信息区域 ===
    stats_container = tk.Frame(main_container, bg=SIDEBAR_BG, relief="solid", bd=1)
    stats_container.pack(fill=tk.X, pady=(0, 15))

    # 获取统计数据
    stats = plan_data.get('statistics', {})
    total_images = stats.get('total_images', 0)
    to_process = stats.get('to_process', 0)
    qualified = stats.get('qualified', 0)

    # 创建统计信息行
    stats_info = tk.Frame(stats_container, bg=SIDEBAR_BG)
    stats_info.pack(fill=tk.X, padx=15, pady=10)

    # 统计信息文本 - 更清晰的显示
    stats_text = f"📊 总计: {total_images} 张  |  ⚡ 待处理: {to_process} 张  |  ✅ 已达标: {qualified} 张"
    stats_label = tk.Label(stats_info, text=stats_text, 
                          font=(FONT_FAMILY, 11, "bold"), 
                          bg=SIDEBAR_BG, fg=TEXT_COLOR)
    stats_label.pack()

    # === 处理计划表格卡片 ===
    table_container = tk.Frame(main_container, bg=CARD_BG, relief="solid", bd=1)
    table_container.pack(fill=tk.X, pady=(0, 10))  # 改为fill=tk.X，不再expand
    table_container.configure(height=320)  # 设置固定高度，减小以确保按钮完整显示
    table_container.pack_propagate(False)  # 禁止子组件影响容器大小
    
    # 表格标题
    table_header = tk.Frame(table_container, bg=SIDEBAR_BG)
    table_header.pack(fill=tk.X, padx=1, pady=1)
    
    table_title = tk.Label(table_header, text="📋 处理计划详情", 
                          font=(FONT_FAMILY, 12, "bold"), 
                          bg=SIDEBAR_BG, fg=TEXT_COLOR)
    table_title.pack(side=tk.LEFT, padx=15, pady=6)
    
    edit_hint = tk.Label(table_header, text="💡 双击或右键编辑倍数", 
                        font=(FONT_FAMILY, 9), 
                        bg=SIDEBAR_BG, fg=TEXT_SECONDARY)
    edit_hint.pack(side=tk.RIGHT, padx=15, pady=6)
    
    # 表格主体容器 - 固定高度的滚动区域
    table_main = tk.Frame(table_container, bg=CARD_BG)
    table_main.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
    
    # 创建现代化表格 - 设置固定高度，防止内容撑大窗口
    columns = ('filename', 'original_size', 'target_size', 'scale_factor')
    tree = ttk.Treeview(table_main, columns=columns, show='headings', style="Modern.Treeview", height=12)

    # 定义列标题和图标
    headers = [
        ('filename', '📁 文件名'),
        ('original_size', '📐 原始尺寸'),
        ('target_size', '🎯 目标尺寸'),
        ('scale_factor', '🔍 放大倍数')
    ]

    for col, header in headers:
        tree.heading(col, text=header)

    # 设置列宽和对齐
    tree.column('filename', width=280, minwidth=200, anchor=tk.W)
    tree.column('original_size', width=130, minwidth=100, anchor=tk.CENTER)
    tree.column('target_size', width=130, minwidth=100, anchor=tk.CENTER)
    tree.column('scale_factor', width=120, minwidth=100, anchor=tk.CENTER)

    # 智能滚动条 - 直接添加到table_main中
    scrollbar = ttk.Scrollbar(table_main, orient=tk.VERTICAL, command=tree.yview, style="Modern.Vertical.TScrollbar")
    tree.configure(yscrollcommand=scrollbar.set)

    # 布局表格和滚动条
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(2, 0))
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 填充数据
    tree.tag_configure('oddrow', background=SIDEBAR_BG)
    tree.tag_configure('evenrow', background=CARD_BG)
    
    processing_list = modified_plan_data.get('processing_list', [])
    tree_items = {}
    
    # 显示所有数据，不再限制显示数量
    for i, item in enumerate(processing_list):
        filename = os.path.basename(item.get('filename', item.get('image_path', '')))
        original_size = f"{item.get('original_width', 0)}×{item.get('original_height', 0)}"
        target_size = f"{item.get('target_width', 0)}×{item.get('target_height', 0)}"
        scale_factor = f"{item.get('factor', 1)}×"
        
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        tree_item = tree.insert('', 'end', values=(filename, original_size, target_size, scale_factor), tags=(tag,))
        tree_items[tree_item] = i

    # 编辑倍数的现代化弹出窗口
    def edit_scale_factor(item_id):
        """编辑选中项的放大倍数"""
        if item_id not in tree_items:
            return
            
        data_index = tree_items[item_id]
        current_item = processing_list[data_index]
        current_scale = current_item.get('factor', 1)
        
        # 创建现代化编辑对话框
        edit_dialog = tk.Toplevel(dialog)
        edit_dialog.title("✏️ 编辑放大倍数")
        edit_dialog.geometry("380x280")
        edit_dialog.resizable(False, False)
        edit_dialog.transient(dialog)
        edit_dialog.grab_set()
        edit_dialog.configure(bg=BG_COLOR)
        
        # 立即隐藏窗口，避免在左上角显示
        edit_dialog.withdraw()
        
        # 居中显示 - 先更新布局但窗口仍然隐藏
        edit_dialog.update_idletasks()
        x = dialog.winfo_x() + (dialog.winfo_width() // 2) - (edit_dialog.winfo_width() // 2)
        y = dialog.winfo_y() + (dialog.winfo_height() // 2) - (edit_dialog.winfo_height() // 2)
        edit_dialog.geometry(f"+{x}+{y}")
        
        # 设置好位置后再显示窗口，避免移动效果
        edit_dialog.deiconify()
        
        # 主容器
        edit_main = tk.Frame(edit_dialog, bg=BG_COLOR)
        edit_main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        edit_title = tk.Label(edit_main, text="✏️ 编辑放大倍数", 
                             font=(FONT_FAMILY, 14, "bold"), 
                             bg=BG_COLOR, fg=TEXT_COLOR)
        edit_title.pack(pady=(0, 15))
        
        # 文件信息框
        file_info_frame = tk.Frame(edit_main, bg=SIDEBAR_BG, relief="solid", bd=1)
        file_info_frame.pack(fill=tk.X, pady=(0, 15))
        
        filename = os.path.basename(current_item.get('image_path', ''))
        file_label = tk.Label(file_info_frame, text=f"📁 文件: {filename}", 
                             font=(FONT_FAMILY, 10, "bold"), 
                             bg=SIDEBAR_BG, fg=TEXT_COLOR)
        file_label.pack(pady=8)
        
        # 当前尺寸信息
        size_info = tk.Label(file_info_frame, 
                            text=f"📐 当前尺寸: {current_item.get('original_width', 0)}×{current_item.get('original_height', 0)}", 
                            font=(FONT_FAMILY, 9), 
                            bg=SIDEBAR_BG, fg=TEXT_SECONDARY)
        size_info.pack(pady=(0, 8))
        
        # 倍数选择区域
        scale_frame = tk.Frame(edit_main, bg=BG_COLOR)
        scale_frame.pack(fill=tk.X, pady=(0, 15))
        
        scale_label = tk.Label(scale_frame, text="🔍 选择放大倍数:", 
                              font=(FONT_FAMILY, 11, "bold"), 
                              bg=BG_COLOR, fg=TEXT_COLOR)
        scale_label.pack(pady=(0, 8))
        
        scale_var = tk.StringVar(value=str(current_scale))
        scale_combo = ttk.Combobox(scale_frame, textvariable=scale_var, 
                                  values=['1', '2', '4', '8', '16'], 
                                  width=15, font=(FONT_FAMILY, 11),
                                  justify='center')
        scale_combo.pack()
        scale_combo.focus_set()
        
        # 预览信息
        # preview_frame = tk.Frame(edit_main, bg=SIDEBAR_BG, relief="solid", bd=1)
        # preview_frame.pack(fill=tk.X, pady=(0, 20))
        
        # preview_label = tk.Label(preview_frame, text="", 
        #                         font=(FONT_FAMILY, 10, "bold"), 
        #                         bg=SIDEBAR_BG, fg=PRIMARY_COLOR)
        # preview_label.pack(pady=8)
        
        def update_preview(*args):
            # 预览功能暂时禁用，减少界面复杂度
            pass
            # try:
            #     new_scale = float(scale_var.get())
            #     original_width = current_item.get('original_width', 0)
            #     original_height = current_item.get('original_height', 0)
            #     new_width = int(original_width * new_scale)
            #     new_height = int(original_height * new_scale)
            #     preview_label.config(text=f"🎯 预览尺寸: {new_width}×{new_height}")
            # except:
            #     preview_label.config(text="")
        
        scale_var.trace('w', update_preview)
        update_preview()
        
        # 按钮区域
        btn_frame = tk.Frame(edit_main, bg=BG_COLOR)
        btn_frame.pack(fill=tk.X)
        
        def save_changes():
            try:
                new_scale = float(scale_var.get())
                if new_scale <= 0:
                    raise ValueError("倍数必须大于0")
                
                # 更新数据
                current_item['factor'] = new_scale
                
                # 重新计算目标尺寸
                original_width = current_item.get('original_width', 0)
                original_height = current_item.get('original_height', 0)
                new_target_width = int(original_width * new_scale)
                new_target_height = int(original_height * new_scale)
                
                current_item['target_width'] = new_target_width
                current_item['target_height'] = new_target_height
                
                # 更新树形控件显示
                new_target_size = f"{new_target_width}×{new_target_height}"
                new_scale_text = f"{new_scale}×"
                
                current_values = list(tree.item(item_id, 'values'))
                current_values[2] = new_target_size
                current_values[3] = new_scale_text
                tree.item(item_id, values=current_values)
                
                edit_dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("输入错误", f"请输入有效的数字: {str(e)}", parent=edit_dialog)
        
        def cancel_edit():
            edit_dialog.destroy()
        
        # 现代化按钮 - 与主程序风格一致
        def create_edit_button(parent, text, command, bg_color, hover_color, is_primary=False):
            btn = tk.Button(parent, text=text, command=command,
                           bg=bg_color, fg='white',
                           font=(FONT_FAMILY, 10, "bold" if is_primary else "normal"),
                           padx=25, pady=10, relief=tk.FLAT,
                           cursor='hand2', bd=0, highlightthickness=0)
            
            def on_enter(e):
                btn.config(bg=hover_color)
            def on_leave(e):
                btn.config(bg=bg_color)
            
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            return btn
        
        # 取消按钮 - 左侧
        cancel_btn = create_edit_button(btn_frame, "❌ 取消", cancel_edit, 
                                       CANCEL_COLOR, CANCEL_HOVER)
        cancel_btn.pack(side=tk.LEFT)
        
        # 保存按钮 - 右侧
        save_btn = create_edit_button(btn_frame, "💾 保存", save_changes, 
                                     SUCCESS_COLOR, SUCCESS_HOVER, True)
        save_btn.pack(side=tk.RIGHT)
        
        # 绑定快捷键
        edit_dialog.bind('<Return>', lambda e: save_changes())
        edit_dialog.bind('<Escape>', lambda e: cancel_edit())

    # 绑定双击和右键事件
    def on_tree_double_click(event):
        item_id = tree.selection()[0] if tree.selection() else None
        if item_id:
            edit_scale_factor(item_id)
    
    tree.bind('<Double-1>', on_tree_double_click)

    def show_context_menu(event):
        item_id = tree.identify_row(event.y)
        if item_id:
            tree.selection_set(item_id)
            context_menu = tk.Menu(dialog, tearoff=0, 
                                 bg=CARD_BG, fg=TEXT_COLOR,
                                 font=(FONT_FAMILY, 9))
            context_menu.add_command(label="✏️ 编辑倍数", 
                                   command=lambda: edit_scale_factor(item_id))
            context_menu.tk_popup(event.x_root, event.y_root)
    
    tree.bind('<Button-3>', show_context_menu)

    # === 底部按钮区域 - 固定在底部，确保完全可见 ===
    bottom_frame = tk.Frame(main_container, bg=BG_COLOR, height=80)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
    bottom_frame.pack_propagate(False)
    
    # 按钮容器 - 居中显示
    button_container = tk.Frame(bottom_frame, bg=BG_COLOR)
    button_container.pack(anchor=tk.CENTER, pady=20)
    
    # 现代化按钮样式 - 与主程序保持一致
    def create_action_button(parent, text, command, bg_color, hover_color, is_primary=False):
        btn = tk.Button(parent, text=text, command=command,
                       bg=bg_color, fg='white',
                       font=(FONT_FAMILY, 11, "bold"),
                       padx=30, pady=12, relief=tk.FLAT,
                       cursor='hand2', bd=0, highlightthickness=0)
        
        def on_enter(e):
            btn.config(bg=hover_color)
        def on_leave(e):
            btn.config(bg=bg_color)
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn
    
    # 取消按钮 - 左侧（按用户要求）
    cancel_btn = create_action_button(button_container, "❌ 取消", dialog.destroy, 
                                     CANCEL_COLOR, CANCEL_HOVER)
    cancel_btn.pack(side=tk.LEFT, padx=(0, 20))
    
    # 开始处理按钮 - 右侧，更醒目
    def on_start():
        dialog.destroy()
        if start_callback:
            start_callback(modified_plan_data)
    
    start_btn = create_action_button(button_container, "🚀 开始处理", on_start, 
                                    SUCCESS_COLOR, SUCCESS_HOVER, True)
    start_btn.pack(side=tk.LEFT)
    start_btn.focus_set()

    # 绑定ESC键关闭对话框
    dialog.bind('<Escape>', lambda e: dialog.destroy())

    # 立即隐藏窗口，避免在左上角显示
    dialog.withdraw()

    # 居中显示 - 先更新布局但窗口仍然隐藏
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")
    
    # 设置好位置后再显示窗口，避免移动效果
    dialog.deiconify()

    return dialog