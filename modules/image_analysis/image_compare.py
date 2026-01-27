#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : image_compare.py 
@Author  : Shawn
@Date    : 2026/1/23 9:34 
@Info    : Description of this file
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from PIL import Image, ImageTk
import subprocess
import platform
import time


class ZoomableImage:
    """可缩放和拖动的图片类"""

    def __init__(self, canvas, image_path, select_callback=None, right_click_callback=None):
        self.canvas = canvas
        self.image_path = image_path
        self.select_callback = select_callback  # 单击选择回调函数
        self.right_click_callback = right_click_callback  # 右键点击回调函数
        self.image = None
        self.image_tk = None
        self.image_id = None
        self.info_text = ""  # 信息文本

        # 缩放和拖动相关变量
        self.scale = 1.0
        self.min_scale = 0.1
        self.max_scale = 5.0
        self.offset_x = 0
        self.offset_y = 0
        self.last_x = 0
        self.last_y = 0
        self.is_dragging = False

        # 单击相关变量
        self.click_time = 0
        self.click_threshold = 0.3  # 单击最大持续时间（秒）
        self.drag_threshold = 5  # 拖动最小距离（像素）
        self.start_x = 0
        self.start_y = 0

        # 信息标签相关
        self.info_label_id = None
        self.info_bg_id = None

        # 加载图片
        self.load_image()

        # 绑定事件
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)  # Windows
        self.canvas.bind("<Button-4>", self.on_mousewheel)  # Linux向上滚动
        self.canvas.bind("<Button-5>", self.on_mousewheel)  # Linux向下滚动
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.canvas.bind("<Button-3>", self.on_right_click)  # 右键点击

    def load_image(self):
        """加载图片"""
        try:
            self.image = Image.open(self.image_path)
            self.display_image()
        except Exception as e:
            print(f"无法加载图片: {e}")

    def set_info_text(self, text):
        """设置要显示的信息文本"""
        self.info_text = str(text)
        self.display_image()

    def display_image(self):
        """显示图片"""
        if self.image:
            # 计算适合画布的大小
            canvas_width = self.canvas.winfo_width() - 20
            canvas_height = self.canvas.winfo_height() - 60  # 减去按钮和信息标签的高度

            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width = 400
                canvas_height = 300

            # 保持宽高比缩放
            img_width, img_height = self.image.size
            ratio = min(canvas_width / img_width, canvas_height / img_height)
            display_width = int(img_width * ratio * self.scale)
            display_height = int(img_height * ratio * self.scale)

            # 缩放图片
            resized_img = self.image.resize((display_width, display_height), Image.Resampling.LANCZOS)
            self.image_tk = ImageTk.PhotoImage(resized_img)

            # 清除旧图片并绘制新图片
            if self.image_id:
                self.canvas.delete(self.image_id)
            if self.info_label_id:
                self.canvas.delete(self.info_label_id)
            if self.info_bg_id:
                self.canvas.delete(self.info_bg_id)

            # 计算图片位置（居中显示，但考虑到顶部有信息标签）
            x = canvas_width // 2 + self.offset_x
            y = (canvas_height // 2) + 30 + self.offset_y  # 向下偏移30像素给信息标签留空间

            self.image_id = self.canvas.create_image(x, y, image=self.image_tk, anchor=tk.CENTER)

            # 在图片上显示信息文本（悬浮层）
            if self.info_text:
                self.display_info_overlay(x, y, display_width, display_height)

    def display_info_overlay(self, img_x, img_y, img_width, img_height):
        """在图片上显示信息悬浮层"""
        # 计算信息框的位置（左上角）
        overlay_x = img_x - img_width // 2 + 20  # 距离左边缘20像素
        overlay_y = img_y - img_height // 2 + 20  # 距离上边缘20像素

        # 确定文本颜色（根据OK/NG）
        if self.info_text.upper() == "OK":
            text_color = "green"
        elif self.info_text.upper() == "NG":
            text_color = "red"
        else:
            text_color = "black"

        # 计算文本大小
        font_size = max(36, int(48 * self.scale))  # 基础36像素，随缩放增大

        # 创建半透明背景框
        text_length = len(self.info_text)
        bg_width = max(120, text_length * font_size // 2)  # 根据文本长度调整宽度
        bg_height = font_size + 20

        # 绘制背景框
        self.info_bg_id = self.canvas.create_rectangle(
            overlay_x, overlay_y,
            overlay_x + bg_width, overlay_y + bg_height,
            fill="white", stipple="gray50",  # 半透明效果
            width=2, outline="black"
        )

        # 将背景框置于图片上方
        self.canvas.tag_raise(self.info_bg_id, self.image_id)

        # 绘制文本
        self.info_label_id = self.canvas.create_text(
            overlay_x + bg_width // 2,
            overlay_y + bg_height // 2,
            text=self.info_text,
            fill=text_color,
            font=("Arial", font_size, "bold"),
            anchor=tk.CENTER
        )

        # 将文本置于背景框上方
        self.canvas.tag_raise(self.info_label_id, self.info_bg_id)

    def on_mousewheel(self, event):
        """鼠标滚轮缩放"""
        # 计算缩放比例
        scale_factor = 1.1
        if event.delta < 0 or event.num == 5:  # 向下滚动
            scale_factor = 0.9

        # 获取鼠标位置
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # 计算新的缩放比例
        new_scale = self.scale * scale_factor
        if self.min_scale <= new_scale <= self.max_scale:
            # 计算基于鼠标位置的偏移
            self.offset_x = (x - (self.canvas.winfo_width() // 2)) * (1 - scale_factor) + self.offset_x * scale_factor
            self.offset_y = (y - (self.canvas.winfo_height() // 2)) * (1 - scale_factor) + self.offset_y * scale_factor
            self.scale = new_scale
            self.display_image()

    def on_button_press(self, event):
        """鼠标按下"""
        self.is_dragging = False
        self.start_x = event.x
        self.start_y = event.y
        self.click_time = time.time()
        self.last_x = event.x
        self.last_y = event.y

    def on_mouse_drag(self, event):
        """鼠标拖动"""
        # 计算拖动距离
        dx = abs(event.x - self.start_x)
        dy = abs(event.y - self.start_y)

        # 如果移动距离超过阈值，开始拖动
        if dx > self.drag_threshold or dy > self.drag_threshold:
            self.is_dragging = True

        if self.is_dragging:
            dx = event.x - self.last_x
            dy = event.y - self.last_y
            self.offset_x += dx
            self.offset_y += dy
            self.last_x = event.x
            self.last_y = event.y
            self.display_image()

    def on_button_release(self, event):
        """鼠标释放"""
        # 计算释放时间
        release_time = time.time()
        duration = release_time - self.click_time

        # 计算移动距离
        dx = abs(event.x - self.start_x)
        dy = abs(event.y - self.start_y)

        # 如果是单击（时间短且移动距离小）
        if duration < self.click_threshold and dx < self.drag_threshold and dy < self.drag_threshold and not self.is_dragging:
            if self.select_callback:
                self.select_callback()

        self.is_dragging = False

    def on_right_click(self, event):
        """右键点击"""
        if self.right_click_callback:
            self.right_click_callback(event)

    def reset_view(self):
        """重置视图"""
        self.scale = 1.0
        self.offset_x = 0
        self.offset_y = 0
        self.display_image()


class ImageComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片对比标注工具")
        self.root.geometry("1400x800")

        # 初始化变量
        self.df = None
        self.sg_image_dir = ""
        self.vm_image_dir = ""
        self.excel_path = ""  # 保存Excel文件路径
        self.current_row_index = None
        self.current_seq_num = None
        self.current_image_name = ""  # 当前图片名称
        self.selected_results = {}
        self.notes = {}  # 存储备注信息
        self.filtered_rows = []

        # 创建右键菜单
        self.create_context_menus()

        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 左侧框架（变窄）
        left_frame = ttk.Frame(main_frame, width=400)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 5))
        left_frame.pack_propagate(False)  # 固定宽度

        # 右侧框架（尽可能宽）
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 左侧控件
        ttk.Label(left_frame, text="图片列表", font=("Arial", 10, "bold")).pack(pady=5)

        # 筛选和搜索框架
        filter_search_frame = ttk.Frame(left_frame)
        filter_search_frame.pack(fill=tk.X, padx=5, pady=2)
        filter_search_frame.grid_columnconfigure(1, weight=1)
        filter_search_frame.grid_columnconfigure(3, weight=1)

        # 筛选下拉菜单
        ttk.Label(filter_search_frame, text="筛选:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.filter_var = tk.StringVar(value="is_equal=0")
        filter_options = ["is_equal=0", "is_equal≠0", "全部"]
        filter_combobox = ttk.Combobox(filter_search_frame, textvariable=self.filter_var, values=filter_options, state="readonly", width=10)
        filter_combobox.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        filter_combobox.bind("<<ComboboxSelected>>", self.on_filter_change)

        # 搜索框
        ttk.Label(filter_search_frame, text="搜索:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=3, sticky=tk.EW, padx=(0, 0))
        search_entry.bind("<Return>", self.on_search)

        # 列表控件
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 添加滚动条
        list_scrollbar = ttk.Scrollbar(list_frame)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(list_frame, columns=("seq", "name", "selected"),
                                 show="headings", yscrollcommand=list_scrollbar.set)
        list_scrollbar.config(command=self.tree.yview)

        self.tree.heading("seq", text="序号")
        self.tree.heading("name", text="图片名称")
        self.tree.heading("selected", text="已选")
        self.tree.column("seq", width=50)
        self.tree.column("name", width=250)
        self.tree.column("selected", width=50)

        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_image_select)

        # 按钮框架
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)

        ttk.Button(button_frame, text="导入Excel", command=self.load_excel).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="保存结果", command=self.save_results).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="重置所有视图", command=self.reset_all_views).pack(fill=tk.X, pady=2)

        # 右侧顶部：显示当前图片名称
        self.current_image_frame = ttk.Frame(right_frame)
        self.current_image_frame.pack(fill=tk.X, pady=(0, 5))

        self.current_image_label = ttk.Label(
            self.current_image_frame,
            text="当前图片：",
            font=("Arial", 12, "bold"),
            foreground="blue"
        )
        self.current_image_label.pack(side=tk.LEFT, padx=5)

        # 右侧图片显示框架 - 使用PanedWindow
        right_paned = ttk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
        right_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # SG图片框架
        sg_frame = ttk.LabelFrame(right_paned, text="原图（SG）", padding=5)
        right_paned.add(sg_frame, weight=1)

        # SG图片画布和滚动条
        sg_canvas_frame = ttk.Frame(sg_frame)
        sg_canvas_frame.pack(fill=tk.BOTH, expand=True)

        sg_vscrollbar = ttk.Scrollbar(sg_canvas_frame, orient=tk.VERTICAL)
        sg_hscrollbar = ttk.Scrollbar(sg_canvas_frame, orient=tk.HORIZONTAL)

        self.sg_canvas = tk.Canvas(sg_canvas_frame, bg='white',
                                   yscrollcommand=sg_vscrollbar.set,
                                   xscrollcommand=sg_hscrollbar.set,
                                   cursor="hand2")  # 鼠标悬停时显示手型

        sg_vscrollbar.config(command=self.sg_canvas.yview)
        sg_hscrollbar.config(command=self.sg_canvas.xview)

        self.sg_canvas.grid(row=0, column=0, sticky="nsew")
        sg_vscrollbar.grid(row=0, column=1, sticky="ns")
        sg_hscrollbar.grid(row=1, column=0, sticky="ew")

        sg_canvas_frame.grid_rowconfigure(0, weight=1)
        sg_canvas_frame.grid_columnconfigure(0, weight=1)

        self.sg_zoomable_image = None

        # SG图片点击提示标签
        self.sg_click_hint = tk.Label(sg_canvas_frame,
                                      text="",
                                      bg="white", fg="blue", font=("Arial", 10, "italic"))
        self.sg_click_hint.place(relx=0.5, rely=0.95, anchor=tk.CENTER)

        # VM图片框架
        vm_frame = ttk.LabelFrame(right_paned, text="结果图（VM）", padding=5)
        right_paned.add(vm_frame, weight=1)

        # VM图片画布和滚动条
        vm_canvas_frame = ttk.Frame(vm_frame)
        vm_canvas_frame.pack(fill=tk.BOTH, expand=True)

        vm_vscrollbar = ttk.Scrollbar(vm_canvas_frame, orient=tk.VERTICAL)
        vm_hscrollbar = ttk.Scrollbar(vm_canvas_frame, orient=tk.HORIZONTAL)

        self.vm_canvas = tk.Canvas(vm_canvas_frame, bg='white',
                                   yscrollcommand=vm_vscrollbar.set,
                                   xscrollcommand=vm_hscrollbar.set,
                                   cursor="hand2")  # 鼠标悬停时显示手型

        vm_vscrollbar.config(command=self.vm_canvas.yview)
        vm_hscrollbar.config(command=self.vm_canvas.xview)

        self.vm_canvas.grid(row=0, column=0, sticky="nsew")
        vm_vscrollbar.grid(row=0, column=1, sticky="ns")
        vm_hscrollbar.grid(row=1, column=0, sticky="ew")

        vm_canvas_frame.grid_rowconfigure(0, weight=1)
        vm_canvas_frame.grid_columnconfigure(0, weight=1)

        self.vm_zoomable_image = None

        # VM图片点击提示标签
        self.vm_click_hint = tk.Label(vm_canvas_frame,
                                      text="",
                                      bg="white", fg="blue", font=("Arial", 10, "italic"))
        self.vm_click_hint.place(relx=0.5, rely=0.95, anchor=tk.CENTER)

        # 备注输入区域
        note_frame = ttk.LabelFrame(right_frame, text="备注", padding=5)
        note_frame.pack(fill=tk.X, pady=5)
        
        self.note_text = tk.Text(note_frame, height=4, wrap=tk.WORD)
        self.note_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.note_text.bind("<KeyRelease>", self.on_note_change)

        # 右侧底部控制按钮
        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Button(control_frame, text="重置SG视图",
                   command=self.reset_sg_view).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="重置VM视图",
                   command=self.reset_vm_view).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="自动下一项",
                   command=self.auto_next).pack(side=tk.LEFT, padx=2)

        # 状态栏
        self.status_label = ttk.Label(root, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定画布大小变化事件
        self.sg_canvas.bind("<Configure>", lambda e: self.on_canvas_configure("SG"))
        self.vm_canvas.bind("<Configure>", lambda e: self.on_canvas_configure("VM"))

    def create_context_menus(self):
        """创建右键菜单"""
        # SG图片右键菜单
        self.sg_context_menu = tk.Menu(self.root, tearoff=0)
        self.sg_context_menu.add_command(label="打开图片", command=lambda: self.open_image_external("SG"))
        self.sg_context_menu.add_separator()
        self.sg_context_menu.add_command(label="重置视图", command=self.reset_sg_view)

        # VM图片右键菜单
        self.vm_context_menu = tk.Menu(self.root, tearoff=0)
        self.vm_context_menu.add_command(label="打开图片", command=lambda: self.open_image_external("VM"))
        self.vm_context_menu.add_separator()
        self.vm_context_menu.add_command(label="重置视图", command=self.reset_vm_view)

    def show_sg_context_menu(self, event):
        """显示SG图片右键菜单"""
        if self.sg_zoomable_image and self.sg_zoomable_image.image:
            try:
                self.sg_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.sg_context_menu.grab_release()

    def show_vm_context_menu(self, event):
        """显示VM图片右键菜单"""
        if self.vm_zoomable_image and self.vm_zoomable_image.image:
            try:
                self.vm_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.vm_context_menu.grab_release()

    def update_status(self, text):
        self.status_label.config(text=text)
        self.root.update()

    def update_current_image_label(self):
        """更新当前图片名称标签"""
        if self.current_image_name:
            display_text = f"当前图片：{self.current_image_name}"
            # 如果名称太长，截断显示
            if len(display_text) > 60:
                display_text = display_text[:57] + "..."
            self.current_image_label.config(text=display_text)
        else:
            self.current_image_label.config(text="当前图片：")

    def apply_filter(self):
        """应用筛选条件，生成筛选后的行集并更新列表"""
        if self.df is None:
            return

        # 查找is_equal列
        is_equal_col = "is_equal" if "is_equal" in self.df.columns else self.df.columns[10]

        # 根据筛选条件筛选数据
        filter_value = self.filter_var.get()
        filtered_df = self.df.copy()

        try:
            is_equal_series = pd.to_numeric(self.df[is_equal_col], errors='coerce')
            if filter_value == "is_equal=0":
                filtered_df = self.df[is_equal_series == 0].copy()
            elif filter_value == "is_equal≠0":
                filtered_df = self.df[is_equal_series != 0].copy()
            # "全部"选项不筛选
        except Exception as e:
            print(f"数值筛选失败，尝试字符串筛选: {e}")
            is_equal_str_series = self.df[is_equal_col].astype(str)
            if filter_value == "is_equal=0":
                filtered_df = self.df[is_equal_str_series == "0"].copy()
            elif filter_value == "is_equal≠0":
                filtered_df = self.df[is_equal_str_series != "0"].copy()
            # "全部"选项不筛选
            
        # 确保filtered_df是DataFrame类型
        if not isinstance(filtered_df, pd.DataFrame):
            filtered_df = self.df.copy()

        # 更新filtered_rows
        self.filtered_rows = []
        for idx, row in filtered_df.iterrows():
            seq_num = str(row[0]) if not pd.isna(row[0]) else str(idx + 1)
            # 去除回车符和前后空格
            seq_num = seq_num.strip().replace('\n', '')
            img_name = str(row[6]) if len(row) > 6 and not pd.isna(row[6]) else "未知"
            # 去除回车符和前后空格
            img_name = img_name.strip().replace('\n', '')
            if isinstance(img_name, str) and os.path.sep in img_name:
                img_name = os.path.basename(img_name)
                # 去除basename后的回车符和前后空格
                img_name = img_name.strip().replace('\n', '')
            self.filtered_rows.append((idx, seq_num, img_name))

        # 应用搜索过滤
        self.on_search()

    def on_note_change(self, event=None):
        """备注文本变化事件"""
        if self.current_seq_num is not None:
            note = self.note_text.get("1.0", tk.END).strip()
            self.notes[self.current_seq_num] = note

    def on_filter_change(self, event=None):
        """筛选条件变化事件"""
        if self.df is not None:
            self.apply_filter()
            self.update_status(f"已应用筛选条件：{self.filter_var.get()}")

    def on_search(self, event=None):
        """搜索功能"""
        if self.df is None:
            return

        search_text = self.search_var.get().lower().strip().replace('\n', '')

        # 清空列表
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 重新插入匹配的行
        for idx, seq_num, img_name in self.filtered_rows:
            if search_text in seq_num.lower() or search_text in img_name.lower():
                # 检查是否已有选择结果
                selected = ""
                if seq_num in self.selected_results:
                    selected = self.selected_results[seq_num]

                # 存储原始idx到列表项中，方便后续获取
                self.tree.insert("", tk.END, values=(seq_num, img_name, selected), tags=(str(idx), seq_num))

    def load_excel(self):
        """加载Excel文件"""
        file_path = filedialog.askopenfilename(title="选择Excel文件",
                                               filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.update_status("正在加载Excel...")

        try:
            self.excel_path = file_path  # 保存原文件路径
            self.df = pd.read_excel(file_path, dtype=str)

            self.sg_image_dir = filedialog.askdirectory(title="选择SG图片目录（第6列图片所在目录）")
            if not self.sg_image_dir:
                self.df = None
                self.excel_path = ""
                return

            self.vm_image_dir = filedialog.askdirectory(title="选择VM图片目录（第7列图片所在目录）")
            if not self.vm_image_dir:
                self.df = None
                self.excel_path = ""
                return

            # 确保self.df是DataFrame类型
            if not isinstance(self.df, pd.DataFrame):
                messagebox.showerror("错误", "加载Excel失败：文件格式不正确")
                self.df = None
                self.excel_path = ""
                return

            if len(self.df.columns) < 7:
                messagebox.showerror("错误", "Excel列数不足，需要至少7列")
                self.df = None
                self.excel_path = ""
                return

            # 读取第8列的现有结果到selected_results字典
            self.selected_results = {}
            # 读取第9列的备注数据到notes字典
            self.notes = {}
            
            for idx, row in self.df.iterrows():
                seq_num = str(row[0]) if not pd.isna(row[0]) else str(idx + 1)
                # 去除回车符和前后空格
                seq_num = seq_num.strip().replace('\n', '')
                
                # 读取第12列的选择结果
                if len(self.df.columns) >= 12:
                    result_col_index = 11
                    result_val = str(row[result_col_index]) if not pd.isna(row[result_col_index]) else ""
                    # 去除回车符和前后空格
                    result_val = result_val.strip().replace('\n', '')
                    if result_val in ["SG", "VM"]:
                        self.selected_results[seq_num] = result_val
                
                # 读取第13列的备注
                if len(self.df.columns) >= 13:
                    note_col_index = 12
                    note_val = str(row[note_col_index]) if not pd.isna(row[note_col_index]) else ""
                    # 去除回车符和前后空格
                    note_val = note_val.strip().replace('\n', '')
                    if note_val:
                        self.notes[seq_num] = note_val

            # 应用筛选条件
            self.apply_filter()

            self.update_status(f"已加载Excel，找到{len(self.filtered_rows)}个待处理项")

        except Exception as e:
            messagebox.showerror("错误", f"加载Excel失败: {str(e)}")
            self.update_status("加载失败")
            self.df = None
            self.excel_path = ""

    def get_image_path(self, col_index, row_index):
        """获取图片路径"""
        if self.df is None or row_index is None:
            return None

        try:
            img_name = str(self.df.iloc[row_index, col_index])
            # 去除回车符和前后空格
            img_name = img_name.strip().replace('\n', '')
            if os.path.sep in img_name:
                img_name = os.path.basename(img_name)
                # 去除basename后的回车符和前后空格
                img_name = img_name.strip().replace('\n', '')

            if col_index == 5:
                return os.path.join(self.sg_image_dir, img_name)
            elif col_index == 6:
                return os.path.join(self.vm_image_dir, img_name)
            else:
                return None
        except:
            return None

    def load_zoomable_image(self, canvas, image_path, image_type, info_text=""):
        """加载可缩放的图片"""
        if image_path and os.path.exists(image_path):
            # 根据图片类型设置回调函数
            if image_type == "SG":
                zoom_image = ZoomableImage(canvas, image_path,
                                           select_callback=lambda: self.select_image("SG"),
                                           right_click_callback=self.show_sg_context_menu)
            else:
                zoom_image = ZoomableImage(canvas, image_path,
                                           select_callback=lambda: self.select_image("VM"),
                                           right_click_callback=self.show_vm_context_menu)

            # 设置信息文本
            if info_text:
                zoom_image.set_info_text(info_text)

            return zoom_image
        return None

    def on_canvas_configure(self, image_type):
        """画布大小变化时重新显示图片"""
        if image_type == "SG" and self.sg_zoomable_image:
            self.sg_zoomable_image.display_image()
            # 更新提示标签位置
            self.sg_click_hint.place(relx=0.5, rely=0.95, anchor=tk.CENTER)
        elif image_type == "VM" and self.vm_zoomable_image:
            self.vm_zoomable_image.display_image()
            # 更新提示标签位置
            self.vm_click_hint.place(relx=0.5, rely=0.95, anchor=tk.CENTER)

    def on_image_select(self, event):
        """左侧列表选择事件"""
        selection = self.tree.selection()
        if not selection:
            return

        item = selection[0]
        # 获取标签中的原始idx和seq_num
        tags = self.tree.item(item, "tags")
        if not tags:
            return
        
        self.current_row_index = int(tags[0])  # 原始索引
        self.current_seq_num = tags[1]  # 序号
        
        # 获取图片名称
        img_name = self.tree.item(item, "values")[1]
        self.current_image_name = img_name

        # 更新当前图片名称标签
        self.update_current_image_label()

        # 显示提示标签
        self.sg_click_hint.lift()
        self.vm_click_hint.lift()

        # 获取SG信息文本（第10列，XML结果）
        sg_info = ""
        if len(self.df.columns) > 9:
            sg_info = str(self.df.iloc[self.current_row_index, 9]) if not pd.isna(
                self.df.iloc[self.current_row_index, 9]) else ""

        # 获取VM信息文本（第2列）
        vm_info = ""
        if len(self.df.columns) > 1:
            vm_info = str(self.df.iloc[self.current_row_index, 1]) if not pd.isna(
                self.df.iloc[self.current_row_index, 1]) else ""

        # 清除SG画布内容
        self.sg_canvas.delete("all")
        self.sg_zoomable_image = None  # 清除旧的可缩放图片对象

        # 加载SG图片
        sg_path = self.get_image_path(5, self.current_row_index)
        if sg_path and os.path.exists(sg_path):
            self.sg_zoomable_image = self.load_zoomable_image(self.sg_canvas, sg_path, "SG", sg_info)
            # 确保提示标签在图片上方
            self.sg_click_hint.lift()
        else:
            self.sg_canvas.create_text(self.sg_canvas.winfo_width() // 2,
                                       self.sg_canvas.winfo_height() // 2,
                                       text="图片不存在", fill="red")
            self.sg_click_hint.place_forget()

        # 清除VM画布内容
        self.vm_canvas.delete("all")
        self.vm_zoomable_image = None  # 清除旧的可缩放图片对象

        # 加载VM图片
        vm_path = self.get_image_path(6, self.current_row_index)
        if vm_path and os.path.exists(vm_path):
            self.vm_zoomable_image = self.load_zoomable_image(self.vm_canvas, vm_path, "VM", vm_info)
            # 确保提示标签在图片上方
            self.vm_click_hint.lift()
        else:
            self.vm_canvas.create_text(self.vm_canvas.winfo_width() // 2,
                                       self.vm_canvas.winfo_height() // 2,
                                       text="图片不存在", fill="red")
            self.vm_click_hint.place_forget()

        # 更新高亮显示
        self.update_canvas_highlight()

        # 显示备注
        note = self.notes.get(self.current_seq_num, "")
        self.note_text.delete("1.0", tk.END)
        self.note_text.insert(tk.END, note)

        self.update_status(f"已加载序号 {self.current_seq_num} 的图片（单击图片选择，拖动查看，右键菜单）")

    def update_canvas_highlight(self):
        """更新画布高亮显示"""
        if self.current_seq_num in self.selected_results:
            selected_result = self.selected_results[self.current_seq_num]
            if selected_result == "SG":
                self.sg_canvas.config(highlightbackground="green", highlightthickness=3)
                self.vm_canvas.config(highlightbackground="gray", highlightthickness=1)
            else:
                self.vm_canvas.config(highlightbackground="green", highlightthickness=3)
                self.sg_canvas.config(highlightbackground="gray", highlightthickness=1)
        else:
            self.sg_canvas.config(highlightbackground="gray", highlightthickness=1)
            self.vm_canvas.config(highlightbackground="gray", highlightthickness=1)

    def select_image(self, result_type):
        """选择图片结果（通过单击图片触发）"""
        if self.current_seq_num is None:
            messagebox.showinfo("提示", "请先在左侧选择图片")
            return

        # 保存选择结果
        self.selected_results[self.current_seq_num] = result_type

        # 更新左侧列表
        for item in self.tree.selection():
            self.tree.set(item, "selected", result_type)

        # 更新高亮显示
        self.update_canvas_highlight()

        # 播放选择提示音（可选）
        try:
            self.root.bell()  # 系统提示音
        except:
            pass

        self.update_status(f"已选择：{result_type}（序号 {self.current_seq_num}）")

        # 自动跳转到下一项
        self.auto_next()

    def auto_next(self):
        """自动选择下一项"""
        if not self.tree.selection():
            return

        current_item = self.tree.selection()[0]
        next_item = self.tree.next(current_item)

        if next_item:
            self.tree.selection_set(next_item)
            self.tree.focus(next_item)
            self.tree.see(next_item)
            self.on_image_select(None)

    def open_image_external(self, image_type):
        """用系统默认图片查看器打开图片"""
        if image_type == "SG" and self.sg_zoomable_image:
            image_path = self.sg_zoomable_image.image_path
        elif image_type == "VM" and self.vm_zoomable_image:
            image_path = self.vm_zoomable_image.image_path
        else:
            return

        if not image_path or not os.path.exists(image_path):
            return

        try:
            if platform.system() == "Windows":
                os.startfile(image_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", image_path])
            else:  # Linux
                subprocess.run(["xdg-open", image_path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片：{str(e)}")

    def reset_sg_view(self):
        """重置SG图片视图"""
        if self.sg_zoomable_image:
            self.sg_zoomable_image.reset_view()

    def reset_vm_view(self):
        """重置VM图片视图"""
        if self.vm_zoomable_image:
            self.vm_zoomable_image.reset_view()

    def reset_all_views(self):
        """重置所有图片视图"""
        self.reset_sg_view()
        self.reset_vm_view()

    def save_results(self):
        """保存结果到原Excel文件"""
        if self.df is None:
            messagebox.showinfo("提示", "请先导入Excel文件")
            return

        if not self.selected_results:
            messagebox.showinfo("提示", "没有需要保存的结果")
            return

        # 确保有第12列和第13列
        if len(self.df.columns) < 12:
            self.df["选择结果"] = ""
            result_col_index = 11
        else:
            result_col_index = 11
        
        if len(self.df.columns) < 13:
            self.df["备注"] = ""
            note_col_index = 12
        else:
            note_col_index = 12

        # 备份原数据（可选）
        backup_path = self.excel_path.replace(".xlsx", "_backup.xlsx")
        if not backup_path.endswith("_backup.xlsx"):
            backup_path += "_backup.xlsx"

        # 更新选择结果和备注
        updated_count = 0
        for idx, row in self.df.iterrows():
            seq_num = str(row[0]) if not pd.isna(row[0]) else str(idx + 1)
            # 去除回车符和前后空格，确保与selected_results和notes中的键匹配
            seq_num = seq_num.strip().replace('\n', '')
            
            # 更新选择结果
            if seq_num in self.selected_results:
                old_value = self.df.iloc[idx, result_col_index] if not pd.isna(
                    self.df.iloc[idx, result_col_index]) else ""
                new_value = self.selected_results[seq_num]
                if old_value != new_value:
                    self.df.iloc[idx, result_col_index] = new_value
                    updated_count += 1
            
            # 更新备注
            note = self.notes.get(seq_num, "")
            old_note = self.df.iloc[idx, note_col_index] if not pd.isna(
                self.df.iloc[idx, note_col_index]) else ""
            if old_note != note:
                self.df.iloc[idx, note_col_index] = note
                updated_count += 1

        if updated_count == 0:
            messagebox.showinfo("提示", "没有需要更新的结果（所有结果都已保存）")
            return

        # 确认保存
        if not messagebox.askyesno("确认保存", f"将更新{updated_count}条记录到原文件：\n{self.excel_path}\n\n确认保存？"):
            return

        try:
            # 先创建备份
            try:
                self.df.to_excel(backup_path, index=False)
            except:
                pass  # 备份失败不影响主要保存

            # 保存到原文件
            self.df.to_excel(self.excel_path, index=False)

            messagebox.showinfo("成功", f"已保存到原文件：\n{self.excel_path}\n更新了{updated_count}条记录")
            self.update_status(f"已保存 {updated_count} 个结果到原文件")

        except Exception as e:
            error_msg = f"保存失败: {str(e)}"
            if "Permission denied" in str(e):
                error_msg += "\n\n文件可能被其他程序打开，请关闭后重试。"
            messagebox.showerror("错误", error_msg)
            self.update_status("保存失败")


def start_image_analysis_app():
    root = tk.Tk()
    app = ImageComparisonApp(root)
    root.mainloop()


if __name__ == "__main__":
    start_image_analysis_app()
