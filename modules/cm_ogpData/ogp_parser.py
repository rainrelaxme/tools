#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : ogp_parser.py
@Author  : Shawn
@Date    : 2026/1/8 9:45
@Info    : Description of this file
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import re
import os
from datetime import datetime
import chardet
from tkinter import ttk
import threading
import copy
import pandas as pd

class DataSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OGP数据排序工具 - 批量版")
        self.root.geometry("900x850")

        # 设置默认值
        self.header_lines = tk.IntVar(value=6)
        self.file_paths = []
        self.output_dir = tk.StringVar(value=os.path.expanduser("~"))
        self.result_text = ""
        self.processing = False
        
        # 三次元相关变量
        self.three_d_file_paths = []
        self.three_d_output_dir = tk.StringVar(value=os.path.expanduser("~"))
        self.three_d_processing = False

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 创建Notebook（Tab页容器）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建第一个Tab：OGP数据处理（原有功能）
        self.ogp_frame = tk.Frame(self.notebook)
        self.notebook.add(self.ogp_frame, text="OGP")
        self.create_ogp_widgets()
        
        # 创建第二个Tab：三次元数据处理
        self.three_d_frame = tk.Frame(self.notebook)
        self.notebook.add(self.three_d_frame, text="三次元")
        self.create_three_d_widgets()
        
        # 状态栏
        self.status_bar = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_ogp_widgets(self):
        """创建OGP数据处理Tab的界面"""
        # 创建主框架，使用grid布局
        main_frame = tk.Frame(self.ogp_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # === 左侧区域：文件选择和参数设置 ===
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # 处理模式选择
        # mode_frame = tk.LabelFrame(left_frame, text="处理模式", padx=8, pady=8)
        # mode_frame.pack(fill="x", pady=(0, 5))
        #
        self.process_mode = tk.StringVar(value="summary")  # 默认汇总模式

        # 文件格式选择
        format_frame = tk.LabelFrame(left_frame, text="文件格式", padx=8, pady=8)
        format_frame.pack(fill="x", pady=5)

        self.file_format = tk.StringVar(value="format1")  # 默认自动检测
        # tk.Radiobutton(format_frame, text="自动检测", variable=self.file_format,
        #                value="auto", font=("Arial", 9)).pack(anchor="w", pady=2)
        tk.Radiobutton(format_frame, text="格式1: 多表头空行分隔", variable=self.file_format,
                       value="format1", font=("Arial", 9)).pack(anchor="w", pady=2)
        tk.Radiobutton(format_frame, text="格式2: 单表头无空行", variable=self.file_format,
                       value="format2", font=("Arial", 9)).pack(anchor="w", pady=2)

        # 文件选择区域
        file_frame = tk.LabelFrame(left_frame, text="文件选择", padx=8, pady=8)
        file_frame.pack(fill="x", pady=5)

        # 文件列表显示
        tk.Label(file_frame, text="已选择的文件:", font=("Arial", 9, "bold")).pack(anchor="w")

        # Treeview框架
        tree_frame = tk.Frame(file_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        # 滚动条
        scrollbar = tk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview
        self.file_tree = ttk.Treeview(tree_frame, columns=("序号", "文件名"),
                                      show="headings", yscrollcommand=scrollbar.set, height=6)

        # 设置列
        self.file_tree.heading("序号", text="序号")
        self.file_tree.heading("文件名", text="文件名")

        self.file_tree.column("序号", width=50, anchor="center")
        self.file_tree.column("文件名", width=250)

        self.file_tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar.config(command=self.file_tree.yview)

        # 右键菜单
        self.file_tree.bind("<Button-3>", self.show_file_context_menu)

        # 文件操作按钮（水平排列）
        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.pack(fill="x", pady=5)

        tk.Button(file_buttons_frame, text="添加文件", command=self.browse_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="移除选中", command=self.remove_selected_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="清空列表", command=self.clear_file_list,
                  width=12).pack(side=tk.LEFT, padx=2)

        # 文件统计信息
        self.file_count_label = tk.Label(file_frame, text="已选择 0 个文件", fg="blue", font=("Arial", 9))
        self.file_count_label.pack(anchor="w")

        # 输出文件夹选择区域
        output_frame = tk.LabelFrame(left_frame, text="输出设置", padx=8, pady=8)
        output_frame.pack(fill="x", pady=5)

        # 输出文件夹输入框和按钮
        output_input_frame = tk.Frame(output_frame)
        output_input_frame.pack(fill="x", pady=2)

        tk.Label(output_input_frame, text="文件夹:", width=8).pack(side=tk.LEFT)
        tk.Entry(output_input_frame, textvariable=self.output_dir, width=30).pack(side=tk.LEFT, padx=2, fill="x",
                                                                                  expand=True)
        tk.Button(output_input_frame, text="选择", command=self.browse_output_dir, width=8).pack(side=tk.LEFT, padx=2)

        # 处理选项
        options_frame = tk.Frame(output_frame)
        options_frame.pack(fill="x", pady=5)

        self.create_subfolder = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="创建子文件夹", variable=self.create_subfolder,
                       font=("Arial", 9)).pack(anchor="w")

        self.overwrite_existing = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="覆盖已存在文件", variable=self.overwrite_existing,
                       font=("Arial", 9)).pack(anchor="w")

        # 参数设置区域
        param_frame = tk.LabelFrame(left_frame, text="处理参数", padx=8, pady=8)
        param_frame.pack(fill="x", pady=5)

        param_input_frame = tk.Frame(param_frame)
        param_input_frame.pack(fill="x", pady=5)

        tk.Label(param_input_frame, text="表头行数:", width=10).pack(side=tk.LEFT)
        tk.Entry(param_input_frame, textvariable=self.header_lines, width=10).pack(side=tk.LEFT, padx=5)

        # 处理按钮区域（在左侧框架底部）
        button_frame = tk.Frame(left_frame)
        button_frame.pack(fill="x", pady=10)

        # 开始处理按钮
        self.start_button = tk.Button(button_frame, text="开始批量处理", command=self.start_batch_processing,
                                      bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                      height=2, state=tk.DISABLED)
        self.start_button.pack(fill="x", pady=2)

        # 停止处理按钮
        self.stop_button = tk.Button(button_frame, text="停止处理", command=self.stop_processing,
                                     bg="#FF9800", fg="white", height=1, state=tk.DISABLED)
        self.stop_button.pack(fill="x", pady=2)

        # 进度条区域
        progress_frame = tk.LabelFrame(left_frame, text="处理进度", padx=8, pady=8)
        progress_frame.pack(fill="x", pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=5)

        self.progress_label = tk.Label(progress_frame, text="就绪", fg="green", font=("Arial", 9))
        self.progress_label.pack()

        # === 右侧区域：结果显示 ===
        right_frame = tk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # 结果显示区域
        result_frame = tk.LabelFrame(right_frame, text="处理结果", padx=8, pady=8)
        result_frame.pack(fill="both", expand=True)

        # 使用ScrolledText显示结果
        self.result_text_area = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Courier New", 9))
        self.result_text_area.pack(fill="both", expand=True)

        # 配置网格权重
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_rowconfigure(0, weight=1)

    def browse_files(self):
        """浏览并选择多个文件"""
        filenames = filedialog.askopenfilenames(
            title="选择OGP数据文件",
            filetypes=[("Excel/Text files", "*.xls *.xlsx *.txt *.csv"), ("All files", "*.*")]
        )

        if filenames:
            for filename in filenames:
                if filename not in self.file_paths:
                    self.file_paths.append(filename)
                    # 添加到Treeview
                    index = len(self.file_paths)
                    self.file_tree.insert("", "end", values=(
                        index,
                        os.path.basename(filename)
                    ))

            # 更新统计信息
            self.update_file_count()

            # 自动设置输出文件夹为第一个文件所在目录
            if self.file_paths and (not self.output_dir.get() or self.output_dir.get() == os.path.expanduser("~")):
                input_dir = os.path.dirname(self.file_paths[0])
                self.output_dir.set(input_dir)

            self.status_bar.config(text=f"已添加 {len(filenames)} 个文件")

    def browse_output_dir(self):
        """浏览并选择输出文件夹"""
        dirname = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self.output_dir.get()
        )
        if dirname:
            self.output_dir.set(dirname)
            self.status_bar.config(text=f"输出文件夹: {dirname}")

    def clear_file_list(self):
        """清空文件列表"""
        if self.file_paths:
            if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
                self.file_paths.clear()
                # 清空Treeview
                for item in self.file_tree.get_children():
                    self.file_tree.delete(item)
                self.update_file_count()
                self.status_bar.config(text="已清空文件列表")

    def remove_selected_files(self):
        """移除选中的文件"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要移除的文件")
            return

        # 确认移除
        if messagebox.askyesno("确认", f"确定要移除选中的 {len(selected_items)} 个文件吗？"):
            for item in selected_items:
                values = self.file_tree.item(item, "values")
                if values:
                    # 从文件路径列表中移除
                    filename = values[1]
                    # 需要找到完整路径
                    for file_path in self.file_paths[:]:
                        if os.path.basename(file_path) == filename:
                            self.file_paths.remove(file_path)
                            break
                # 从Treeview中移除
                self.file_tree.delete(item)

            # 重新排序
            self.reorder_file_list()
            self.update_file_count()
            self.status_bar.config(text=f"已移除 {len(selected_items)} 个文件")

    def reorder_file_list(self):
        """重新排序文件列表"""
        # 清空Treeview
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # 重新添加
        for i, filename in enumerate(self.file_paths, 1):
            self.file_tree.insert("", "end", values=(
                i,
                os.path.basename(filename)
            ))

    def update_file_count(self):
        """更新文件计数"""
        count = len(self.file_paths)
        self.file_count_label.config(text=f"已选择 {count} 个文件")

        # 根据文件数量启用/禁用处理按钮
        if count > 0:
            self.start_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.DISABLED)

    def show_file_context_menu(self, event):
        """显示文件右键菜单"""
        # 选择右键点击的项目
        item = self.file_tree.identify_row(event.y)
        if item:
            self.file_tree.selection_set(item)

            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="打开文件", command=lambda: self.open_selected_file())
            menu.add_command(label="打开所在文件夹", command=lambda: self.open_file_folder())
            menu.add_separator()
            menu.add_command(label="从列表中移除", command=lambda: self.remove_selected_files())
            menu.tk_popup(event.x_root, event.y_root)

    def open_selected_file(self):
        """打开选中的文件"""
        selected_items = self.file_tree.selection()
        if selected_items:
            values = self.file_tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in self.file_paths:
                    if os.path.basename(file_path) == filename:
                        try:
                            if os.name == 'nt':
                                os.startfile(file_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', file_path] if os.sys.platform == 'darwin' else ['xdg-open', file_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件:\n{str(e)}")
                        break

    def open_file_folder(self):
        """打开文件所在文件夹"""
        selected_items = self.file_tree.selection()
        if selected_items:
            values = self.file_tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in self.file_paths:
                    if os.path.basename(file_path) == filename:
                        folder_path = os.path.dirname(file_path)
                        try:
                            if os.name == 'nt':
                                os.startfile(folder_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', folder_path] if os.sys.platform == 'darwin' else ['xdg-open', folder_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件夹:\n{str(e)}")
                        break

    def start_batch_processing(self):
        """开始批量处理"""
        if not self.file_paths:
            messagebox.showerror("错误", "请先选择要处理的文件！")
            return

        # 检查输出文件夹
        if not self.output_dir.get():
            messagebox.showerror("错误", "请选择输出文件夹！")
            return

        # 创建输出文件夹（如果需要）
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹:\n{str(e)}")
            return

        # 禁用处理按钮，启用停止按钮
        self.toggle_buttons(processing=True)
        self.processing = True

        # 在新线程中处理
        thread = threading.Thread(target=self.process_files_thread, daemon=True)
        thread.start()

    def stop_processing(self):
        """停止处理"""
        self.processing = False
        self.status_bar.config(text="正在停止处理...")

    def toggle_buttons(self, processing):
        """切换按钮状态"""
        if processing:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.NORMAL if self.file_paths else tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)

    def process_files_thread(self):
        """处理文件的线程函数"""
        total_files = len(self.file_paths)
        success_count = 0
        fail_count = 0

        # 清空结果区域
        self.result_text_area.delete(1.0, tk.END)

        # 显示开始信息
        self.show_result_header(total_files)

        for i, file_path in enumerate(self.file_paths, 1):
            if not self.processing:
                self.append_result("\n\n处理已停止！\n")
                break

            # 更新进度
            progress = (i / total_files) * 100
            self.progress_var.set(progress)

            # 根据处理模式显示不同文本
            mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"
            self.progress_label.config(text=f"{mode_text} {i}/{total_files}")
            self.status_bar.config(text=f"正在处理: {os.path.basename(file_path)}")

            # 处理单个文件
            success, message = self.process_single_file(file_path, i)

            if success:
                success_count += 1
                self.append_result(f"✓ {message}\n")
            else:
                fail_count += 1
                self.append_result(f"✗ {message}\n")

            # 更新进度条颜色
            if fail_count > 0:
                self.progress_bar.config(style="red.Horizontal.TProgressbar")

        # 处理完成
        self.processing = False
        self.toggle_buttons(processing=False)
        self.progress_var.set(100)

        # 显示总结
        self.show_result_summary(success_count, fail_count)

        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"
        self.progress_label.config(text=f"{mode_text}完成")
        self.status_bar.config(text=f"批量处理完成: 成功 {success_count} 个，失败 {fail_count} 个")

    def process_single_file(self, file_path, index):
        """处理单个文件"""
        try:
            # 检测编码
            encoding = detect_encoding(file_path)

            # 读取文件
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            # with open(file_path, 'r', errors='ignore') as f:
                content = f.read()

            # 确定文件格式
            file_format = self.detect_file_format(content)

            # 根据处理模式选择处理方法
            if self.process_mode.get() == "summary":
                # 汇总模式
                if file_format == "format2":
                    result, block_count = self.summarize_format2_data(content)
                else:
                    result, block_count = self.summarize_data(content)
                mode_text = "汇总"
            else:
                # 仅排序模式
                if file_format == "format2":
                    result, block_count = self.sort_format2_data(content)
                else:
                    result, block_count = self.sort_data(content)
                mode_text = "排序"

            # 确定输出路径
            input_filename = os.path.basename(file_path)
            base_name = os.path.splitext(input_filename)[0]
            suffix = os.path.splitext(input_filename)[1]
            if not suffix:
                suffix = ".txt"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            mode_suffix = "_sorted" if self.process_mode.get() == "sort" else "_summarized"

            # 如果需要创建子文件夹
            if self.create_subfolder.get():
                subfolder_name = f"{mode_text}_{timestamp}"
                output_folder = os.path.join(self.output_dir.get(), subfolder_name)
                os.makedirs(output_folder, exist_ok=True)
                output_filename = f"{base_name}{mode_suffix}_{timestamp}{suffix}"
                output_file = os.path.join(output_folder, output_filename)
            else:
                output_folder = self.output_dir.get()
                # 检查是否覆盖
                output_filename = f"{base_name}{mode_suffix}_{timestamp}{suffix}"
                output_file = os.path.join(output_folder, output_filename)

                # 如果文件已存在且不覆盖，添加序号
                if os.path.exists(output_file) and not self.overwrite_existing.get():
                    counter = 1
                    while os.path.exists(output_file):
                        output_filename = f"{base_name}{mode_suffix}_{timestamp}_{counter}{suffix}"
                        output_file = os.path.join(output_folder, output_filename)
                        counter += 1

            # 保存文件
            with open(output_file, 'w', encoding=encoding, errors='ignore') as f:
                f.write(result)

            return True, f"[{index}] {input_filename}: {mode_text}处理完成，识别到 {block_count} 个区块 -> {output_filename}"

        except Exception as e:
            return False, f"[{index}] {os.path.basename(file_path)}: 处理失败 - {str(e)}"

    def detect_file_format(self, content):
        """检测文件格式"""
        user_format = self.file_format.get()

        if user_format != "auto":
            return user_format

        # 自动检测格式
        lines = content.strip().split('\n')

        # 检查是否有空行
        has_empty_lines = any(line.strip() == '' for line in lines)

        if has_empty_lines:
            return "format1"  # 带空行分隔的格式

        # 检查是否是格式2（单表头，通过标签重复区分区块）
        # 统计标签出现的次数
        label_counts = {}
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*\s+|[\d\.]+\s+)')

        for line in lines:
            if data_line_pattern.match(line.strip()):
                parts = line.strip().split('\t')
                if parts:
                    label = parts[0].strip()
                    label_counts[label] = label_counts.get(label, 0) + 1

        # 如果有标签出现多次，可能是格式2
        max_count = max(label_counts.values()) if label_counts else 0

        if max_count > 1:
            # 进一步确认：检查是否有明显的表头行
            # 格式2通常有特定的表头关键词
            header_keywords = ['±êÇ©', '³ß´çÀàÐÍ', '±ê×¼Öµ', 'Êµ²âÖµ', 'ÉÏ¹«²î', 'ÏÂ¹«²î', 'Æ«²îÖµ']
            has_cn_header = any(keyword in content for keyword in header_keywords)

            if has_cn_header:
                return "format2"

        return "format1"  # 默认为格式1

    def show_result_header(self, total_files):
        """显示结果头部信息"""
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"
        format_text = "自动检测" if self.file_format.get() == "auto" else self.file_format.get()

        result_info = f"{'=' * 70}\n"
        result_info += f"批量{mode_text}开始 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        result_info += f"{'=' * 70}\n\n"
        result_info += f"文件总数: {total_files}\n"
        result_info += f"输出文件夹: {self.output_dir.get()}\n"
        result_info += f"表头行数: {self.header_lines.get()}\n"
        result_info += f"处理模式: {mode_text}\n"
        result_info += f"文件格式: {format_text}\n\n"
        result_info += f"{'=' * 70}\n"
        result_info += "处理结果:\n"
        result_info += f"{'=' * 70}\n\n"

        self.result_text_area.insert(1.0, result_info)

    def append_result(self, message):
        """在结果区域追加信息"""
        self.result_text_area.insert(tk.END, message)
        self.result_text_area.see(tk.END)
        self.root.update()

    def show_result_summary(self, success_count, fail_count):
        """显示处理总结"""
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"

        summary = f"\n{'=' * 70}\n"
        summary += f"批量{mode_text}完成 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        summary += f"{'=' * 70}\n\n"
        summary += f"✓ 成功处理: {success_count} 个文件\n"
        summary += f"✗ 处理失败: {fail_count} 个文件\n"
        summary += f"共计: {success_count + fail_count} 个文件\n\n"

        if fail_count == 0:
            summary += f"🎉 所有文件{mode_text}成功！\n"
        else:
            summary += f"⚠️  有 {fail_count} 个文件{mode_text}失败\n"

        summary += f"{'=' * 70}\n"

        self.append_result(summary)

    def sort_data(self, content):
        """仅排序模式：处理文件内容，按区块排序第一列数据（格式1）"""
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块（通过空行分隔）
        for line in lines:
            if line.strip() == '':
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                current_block.append(line)

        if current_block:
            blocks.append(current_block)

        block_count = len(blocks)
        processed_blocks = []

        # 处理每个区块
        for block in blocks:
            processed_block = self.process_block(block, data_line_pattern)
            processed_blocks.append(processed_block)

        # 合并所有区块（用空行分隔）
        result_lines = []
        for i, block in enumerate(processed_blocks):
            if i > 0:
                result_lines.append('')
            result_lines.extend(block)

        return '\n'.join(result_lines), block_count

    def sort_format2_data(self, content):
        """仅排序模式：处理格式2的数据（单表头无空行）"""
        lines = content.strip().split('\n')
        block_count = 1  # 格式2只有一个表头，数据是连续的

        # 分离表头和数据
        header_lines = []
        data_lines = []
        header_end_index = -1

        # 找到表头结束的位置
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*\s+|[\d\.]+\s+)')
        for i, line in enumerate(lines):
            if data_line_pattern.match(line.strip()):
                header_end_index = i
                break
            else:
                header_lines.append(line)

        # 如果没找到数据行，直接返回原始内容
        if header_end_index == -1:
            return content, 0

        # 提取所有数据行
        data_lines = lines[header_end_index:]

        # 按标签排序数据行
        sorted_data_lines = self.sort_data_lines(data_lines, data_line_pattern)

        # 重新组合
        result_lines = header_lines + sorted_data_lines

        return '\n'.join(result_lines), block_count

    def summarize_data(self, content):
        """
        汇总模式：处理格式1的数据
        1. 先排序数据
        2. 调整列顺序并合并实测值
        3. 只保留第一个区块
        """
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块（通过空行分隔）
        for line in lines:
            if line.strip() == '':
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                current_block.append(line)

        if current_block:
            blocks.append(current_block)

        block_count = len(blocks)

        if block_count == 0:
            return content, 0

        # 处理每个区块（先排序）
        processed_blocks = []
        for block in blocks:
            # 先排序数据行
            sorted_block = self.process_block(block, data_line_pattern)
            processed_blocks.append(sorted_block)

        if block_count == 1:
            # 只有一个区块，只需调整列顺序
            result = self.reorder_and_format_block(processed_blocks[0])
            return '\n'.join(result), block_count

        # 多个区块的情况
        # 1. 从所有区块提取数据
        all_data = self.extract_data_from_blocks(processed_blocks)

        # 2. 按照标签排序数据
        sorted_data = self.sort_extracted_data(all_data)

        # 3. 重新构建输出
        result_block = self.rebuild_output_block(processed_blocks[0][:self.header_lines.get()], sorted_data)

        return '\n'.join(result_block), block_count

    def summarize_format2_data(self, content):
        """
        汇总模式：处理格式2的数据（单表头无空行）
        通过标签重复来识别不同的数据块
        """
        lines = content.strip().split('\n')

        # 分离表头和数据
        header_lines = []
        data_lines = []
        # data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*\s+|[\d\.]+\s+)')
        data_line_pattern = re.compile(r'^[A-Za-z0-9_\-+=*.@#\$%^&()\[\]{}\|;:,.<>?/ \t].*')

        for i, line in enumerate(lines):
            if data_line_pattern.match(line.strip()):
                # 从第一个数据行开始
                data_lines = lines[i:]
                header_lines = lines[:i]
                break

        if not data_lines:
            return content, 0

        # 解析数据行，通过标签重复来识别区块
        blocks = self.split_blocks_by_label_repeat(data_lines)
        block_count = len(blocks)

        if block_count == 0:
            return content, 0

        if block_count == 1:
            # 只有一个区块，只需调整列顺序
            result_block = header_lines + self.reorder_and_format_block_format2(blocks[0])
            return '\n'.join(result_block), block_count

        # 多个区块的情况
        # 1. 从所有区块提取数据
        all_data = self.extract_data_from_blocks_format2(blocks)

        # 2. 按照标签排序数据
        sorted_data = self.sort_extracted_data(all_data)

        # 3. 重新构建输出
        result_block = self.rebuild_output_block_format2(header_lines, sorted_data)

        return '\n'.join(result_block), block_count

    def split_blocks_by_label_repeat(self, data_lines):
        """
        通过标签重复来拆分数据块
        返回列表，每个元素是一个区块的数据行
        """
        blocks = []
        current_block = []
        seen_labels = set()
        block_started = False

        for line in data_lines:
            # 解析标签
            columns = line.strip().split('\t')
            if not columns:
                continue

            label = columns[0].strip()

            # 如果是第一次看到这个标签，或者标签重复出现
            if label in seen_labels and block_started:
                # 标签重复，开始新的区块
                blocks.append(current_block)
                current_block = []
                seen_labels = {label}
                current_block.append(line)
            else:
                # 继续当前区块
                current_block.append(line)
                seen_labels.add(label)
                block_started = True

        # 添加最后一个区块
        if current_block:
            blocks.append(current_block)

        return blocks

    def extract_data_from_blocks_format2(self, blocks):
        """
        从格式2的所有区块中提取数据
        返回字典：{标签: {区块索引: 数据行}}
        """
        data_dict = {}

        for block_idx, block in enumerate(blocks):
            for line in block:
                columns = self.parse_data_line_format2(line)
                if columns and len(columns) >= 6:
                    label = columns[0]  # 第一列是标签
                    if label not in data_dict:
                        data_dict[label] = {}
                    data_dict[label][block_idx] = columns

        return data_dict

    def parse_data_line_format2(self, line):
        """解析格式2的数据行"""
        # 格式2的数据行通常是制表符分隔的
        columns = line.strip().split('\t')
        if len(columns) >= 6:
            return columns
        return None

    def reorder_and_format_block_format2(self, block):
        """调整格式2单个区块的列顺序"""
        formatted_lines = []

        for line in block:
            columns = self.parse_data_line_format2(line)
            if columns and len(columns) >= 6:
                # 格式2的列顺序：标签, 标准值, 上公差, 下公差, 偏差值, 实测值
                # 我们需要重新排序为：标签, 类型(?), 标准值, 上公差, 下公差, 实测值
                label = columns[0]
                nominal = columns[1]
                upper_tol = columns[2]
                lower_tol = columns[3]
                deviation = columns[4]  # 偏差值，可能不需要
                measured = columns[5]

                # 格式化数字
                nominal_formatted = self.format_number(nominal)
                upper_tol_formatted = self.format_number(upper_tol)
                lower_tol_formatted = self.format_number(lower_tol)
                measured_formatted = self.format_number(measured)

                # 构建格式化行
                # 注意：格式2没有明确的"类型"列，我们可以用默认值或留空
                dim_type = "D"  # 默认类型
                formatted_line = f"{label}\t{dim_type}\t{nominal_formatted}\t{upper_tol_formatted}\t{lower_tol_formatted}\t{measured_formatted}"
                formatted_lines.append(formatted_line)
            else:
                formatted_lines.append(line)

        return formatted_lines

    def rebuild_output_block_format2(self, header, sorted_data):
        """重新构建格式2的输出区块"""
        result = header.copy()

        # 确定最大实测值数量（即最大区块数）
        max_block_count = 0
        for sort_key, label, block_data in sorted_data:
            block_count = len(block_data)
            if block_count > max_block_count:
                max_block_count = block_count

        # 修改表头行
        if result:
            # 修改最后一行表头
            last_header = result[-1]
            header_parts = last_header.strip().split('\t')
            if len(header_parts) >= 6:
                # 构建新的表头，添加实测值#1, 实测值#2等
                formatted_title = "标签\t尺寸类型\t标准值\t上公差\t下公差"
                for i in range(1, max_block_count + 1):
                    formatted_title += f"\t实测值#{i}"
                result[-1] = formatted_title

        for sort_key, label, block_data in sorted_data:
            # 获取第一个区块的数据作为基础
            if 0 in block_data:
                base_columns = block_data[0]

                # 提取基础信息
                dim_type = base_columns[1]
                nominal = self.format_number(base_columns[2])
                upper_tol = self.format_number(base_columns[4])
                lower_tol = self.format_number(base_columns[5])

                # 收集所有区块的实测值
                measurements = []
                for block_idx in sorted(block_data.keys()):
                    block_columns = block_data[block_idx]
                    if len(block_columns) >= 6:
                        measurements.append(self.format_number(block_columns[3]))  # 实测值在第6列

                # 构建输出行
                formatted_line = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"

                # 添加所有实测值
                for measurement in measurements:
                    formatted_line += f"\t{measurement}"

                result.append(formatted_line)

        return result

    def process_block(self, block, data_line_pattern):
        """处理单个区块（排序）"""
        data_start = -1
        header_lines = self.header_lines.get()

        if len(block) > header_lines:
            for i in range(header_lines, len(block)):
                if data_line_pattern.match(block[i]):
                    data_start = i
                    break

        if data_start == -1:
            for i, line in enumerate(block):
                if data_line_pattern.match(line):
                    data_start = i
                    break

        if data_start == -1:
            return block

        header = block[:data_start]
        data_lines = block[data_start:]

        sorted_data_lines = self.sort_data_lines(data_lines, data_line_pattern)

        return header + sorted_data_lines

    def sort_data_lines(self, data_lines, data_line_pattern):
        """对数据行进行排序"""
        data_with_keys = []

        for i, line in enumerate(data_lines):
            match = data_line_pattern.match(line)
            if match:
                first_col = match.group(1).strip()
                
                # 初始化排序键
                sort_key = None
                
                try:
                    # 检查是否以数字开头
                    if first_col[0].isdigit():
                        # 数字开头的标签
                        if '*' in first_col:
                            # 处理带星号的标签，如 "15*1", "15*2"
                            num_parts = first_col.split('*')
                            main_num = int(num_parts[0])
                            sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                            sort_key = (0, '', main_num, sub_num, i)
                        else:
                            # 处理纯数字标签，如 "1", "2", "11"
                            main_num = int(first_col)
                            sort_key = (0, '', main_num, 0, i)
                    else:
                        # 字母开头的标签
                        # 提取字母部分和数字部分
                        alpha_part = ''
                        num_part = ''
                        star_part = ''
                        
                        # 找到第一个数字的位置
                        for j, char in enumerate(first_col):
                            if char.isdigit():
                                alpha_part = first_col[:j]
                                remaining = first_col[j:]
                                # 处理可能的星号
                                if '*' in remaining:
                                    num_star_parts = remaining.split('*')
                                    num_part = num_star_parts[0]
                                    star_part = num_star_parts[1] if len(num_star_parts) > 1 else '0'
                                else:
                                    num_part = remaining
                                    star_part = '0'
                                break
                        
                        # 转换数字部分为整数
                        try:
                            num_val = int(num_part)
                            star_val = int(star_part)
                        except ValueError:
                            num_val = 0
                            star_val = 0
                        
                        # 字母开头的排序键
                        sort_key = (1, alpha_part, num_val, star_val, i)
                except (ValueError, IndexError):
                    # 如果解析失败，放到最后
                    sort_key = (2, '', 0, 0, i)
                
                data_with_keys.append((sort_key, line))
            else:
                data_with_keys.append(((2, '', 0, 0, i), line))

        data_with_keys.sort(key=lambda x: x[0])

        return [line for _, line in data_with_keys]

    def extract_data_from_blocks(self, blocks):
        """
        从所有区块中提取数据
        返回字典：{标签: {区块索引: 数据行}}
        """
        header_lines = self.header_lines.get()
        data_dict = {}

        for block_idx, block in enumerate(blocks):
            if len(block) > header_lines:
                for line in block[header_lines:]:
                    # 解析数据行
                    columns = self.parse_data_line(line)
                    if columns:
                        label = columns[0]  # 第一列是标签
                        if label not in data_dict:
                            data_dict[label] = {}
                        data_dict[label][block_idx] = columns

        return data_dict

    def parse_data_line(self, line):
        """解析数据行，返回列列表"""
        # 使用正则表达式匹配行中的列
        # 匹配模式：标签 类型 值1 值2 值3 值4 值5 值6 值7
        pattern1 = r'^\s*([^\s]+)\s+([^\s]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)\s+([-\d\.]+)'
        match = re.match(pattern1, line.strip())

        if match:
            return list(match.groups())

        # 尝试用制表符分割
        columns = line.strip().split('\t')
        if len(columns) >= 6:
            return columns

        return None

    def sort_extracted_data(self, data_dict):
        """对提取的数据进行排序"""
        # 将数据转换为列表以便排序
        data_list = []
        for label, block_data in data_dict.items():
            # 去除首尾空格
            label_stripped = label.strip()

            # 初始化排序键
            sort_key = None

            # 尝试解析标签
            try:
                # 分割主要部分和后缀
                parts = label_stripped.split(None, 1)
                main_part = parts[0]
                suffix = parts[1] if len(parts) > 1 else ""

                # 检查是否以数字开头
                if main_part[0].isdigit():
                    # 数字开头的标签
                    if '*' in main_part:
                        # 处理带星号的标签，如 "15*1", "15*2"
                        num_parts = main_part.split('*')
                        main_num = int(num_parts[0])
                        sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                        sort_key = (0, '', main_num, sub_num, suffix, label_stripped)
                    else:
                        # 处理纯数字标签，如 "1", "2", "11"
                        main_num = int(main_part)
                        sort_key = (0, '', main_num, 0, suffix, label_stripped)
                else:
                    # 字母开头的标签
                    # 提取字母部分和数字部分
                    alpha_part = ''
                    num_part = ''
                    star_part = ''
                    
                    # 找到第一个数字的位置
                    for i, char in enumerate(main_part):
                        if char.isdigit():
                            alpha_part = main_part[:i]
                            remaining = main_part[i:]
                            # 处理可能的星号
                            if '*' in remaining:
                                num_star_parts = remaining.split('*')
                                num_part = num_star_parts[0]
                                star_part = num_star_parts[1] if len(num_star_parts) > 1 else '0'
                            else:
                                num_part = remaining
                                star_part = '0'
                            break
                    
                    # 转换数字部分为整数
                    try:
                        num_val = int(num_part)
                        star_val = int(star_part)
                    except ValueError:
                        num_val = 0
                        star_val = 0
                    
                    # 字母开头的排序键：(1, 字母部分, 数字部分, 星号部分, 后缀, 原始标签)
                    sort_key = (1, alpha_part, num_val, star_val, suffix, label_stripped)

            except (ValueError, IndexError):
                # 如果解析失败，放到最后
                sort_key = (2, '', 0, 0, '', label_stripped)

            data_list.append((sort_key, label, block_data))

        # 按标签排序
        data_list.sort(key=lambda x: x[0])

        return data_list

    def reorder_and_format_block(self, block):
        """调整单个区块的列顺序"""
        header_lines = self.header_lines.get()

        if len(block) <= header_lines:
            return block

        header = block[:header_lines]
        data_lines = block[header_lines:]

        formatted_lines = []

        for line in data_lines:
            columns = self.parse_data_line(line)
            if columns and len(columns) >= 9:
                # 根据示例输出文件，列顺序应该是：
                # 标签, 类型, 标准值, 上公差, 下公差, 偏差?, 0?, 百分比?, 实测值1
                label = columns[0]
                dim_type = columns[1]
                nominal = columns[2]
                measured = columns[3]
                upper_tol = columns[4]
                lower_tol = columns[5]
                # 后面几列可能需要调整
                other_cols = columns[6]  # 第6,7,8列

                # 重新格式化行
                # 注意：根据示例，标准值可能需要去除多余的0
                nominal_formatted = self.format_number(nominal)

                formatted_line = f"{label}\t{dim_type}\t{nominal_formatted}\t{upper_tol}\t{lower_tol}\t{measured}"
                formatted_lines.append(formatted_line)
            else:
                formatted_lines.append(line)

        return header + formatted_lines

    def rebuild_output_block(self, header, sorted_data):
        """重新构建输出区块"""
        result = header.copy()
        
        # 确定最大实测值数量（即最大区块数）
        max_block_count = 0
        for sort_key, label, block_data in sorted_data:
            block_count = len(block_data)
            if block_count > max_block_count:
                max_block_count = block_count
        
        # 排序标题行
        title = result[-1].split('\t')
        label = title[0]
        dim_type = title[1]
        nominal = title[2]
        upper_tol = title[4]
        lower_tol = title[5]
        
        # 构建新的表头，添加实测值#1, 实测值#2等
        formatted_title = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"
        for i in range(1, max_block_count + 1):
            formatted_title += f"\t实测值#{i}"
        result[-1] = formatted_title

        for sort_key, label, block_data in sorted_data:
            # 获取第一个区块的数据作为基础
            if 0 in block_data:
                base_columns = block_data[0]

                # 提取基础信息
                dim_type = base_columns[1]
                nominal = self.format_number(base_columns[2])
                upper_tol = base_columns[4]
                lower_tol = base_columns[5]

                # 收集所有区块的实测值
                measurements = []
                for block_idx in sorted(block_data.keys()):
                    block_columns = block_data[block_idx]
                    if len(block_columns) > 3:
                        measurements.append(block_columns[3])

                # 构建输出行
                formatted_line = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}"

                # 添加所有实测值
                for measurement in measurements:
                    formatted_line += f"\t{measurement}"

                result.append(formatted_line)

        return result

    def format_number(self, num_str):
        """格式化数字，去除多余的0"""
        try:
            # 尝试转换为浮点数
            num = float(num_str)
            # 如果是整数，显示为整数形式
            if num.is_integer():
                return str(int(num))
            else:
                # 去除末尾的0
                formatted = str(num).rstrip('0').rstrip('.')
                return formatted
        except ValueError:
            return num_str

    def create_three_d_widgets(self):
        """创建三次元数据处理Tab的界面"""
        # 创建主框架，使用grid布局
        main_frame = tk.Frame(self.three_d_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # === 左侧区域：文件选择和参数设置 ===
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # 文件选择区域
        file_frame = tk.LabelFrame(left_frame, text="Excel文件选择", padx=8, pady=8)
        file_frame.pack(fill="x", pady=5)

        # 文件列表显示
        tk.Label(file_frame, text="已选择的Excel文件:", font=("Arial", 9, "bold")).pack(anchor="w")

        # Treeview框架
        tree_frame = tk.Frame(file_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        # 滚动条
        scrollbar_3d = tk.Scrollbar(tree_frame)
        scrollbar_3d.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview
        self.three_d_file_tree = ttk.Treeview(tree_frame, columns=("序号", "文件名"),
                                              show="headings", yscrollcommand=scrollbar_3d.set, height=8)

        # 设置列
        self.three_d_file_tree.heading("序号", text="序号")
        self.three_d_file_tree.heading("文件名", text="文件名")

        self.three_d_file_tree.column("序号", width=50, anchor="center")
        self.three_d_file_tree.column("文件名", width=250)

        self.three_d_file_tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar_3d.config(command=self.three_d_file_tree.yview)

        # 右键菜单
        self.three_d_file_tree.bind("<Button-3>", self.show_three_d_file_context_menu)

        # 文件操作按钮（水平排列）
        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.pack(fill="x", pady=5)

        tk.Button(file_buttons_frame, text="添加Excel文件", command=self.browse_three_d_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="移除选中", command=self.remove_selected_three_d_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="清空列表", command=self.clear_three_d_file_list,
                  width=12).pack(side=tk.LEFT, padx=2)

        # 文件统计信息
        self.three_d_file_count_label = tk.Label(file_frame, text="已选择 0 个文件", fg="blue", font=("Arial", 9))
        self.three_d_file_count_label.pack(anchor="w")

        # 输出文件夹选择区域
        output_frame = tk.LabelFrame(left_frame, text="输出设置", padx=8, pady=8)
        output_frame.pack(fill="x", pady=5)

        # 输出文件夹输入框和按钮
        output_input_frame = tk.Frame(output_frame)
        output_input_frame.pack(fill="x", pady=2)

        tk.Label(output_input_frame, text="文件夹:", width=8).pack(side=tk.LEFT)
        tk.Entry(output_input_frame, textvariable=self.three_d_output_dir, width=30).pack(side=tk.LEFT, padx=2, fill="x",
                                                                                          expand=True)
        tk.Button(output_input_frame, text="选择", command=self.browse_three_d_output_dir, width=8).pack(side=tk.LEFT, padx=2)

        # 处理按钮区域
        button_frame = tk.Frame(left_frame)
        button_frame.pack(fill="x", pady=10)

        # 开始处理按钮
        self.three_d_start_button = tk.Button(button_frame, text="开始汇总", command=self.start_three_d_processing,
                                              bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                              height=2, state=tk.DISABLED)
        self.three_d_start_button.pack(fill="x", pady=2)

        # 停止处理按钮
        self.three_d_stop_button = tk.Button(button_frame, text="停止处理", command=self.stop_three_d_processing,
                                             bg="#FF9800", fg="white", height=1, state=tk.DISABLED)
        self.three_d_stop_button.pack(fill="x", pady=2)

        # 进度条区域
        progress_frame = tk.LabelFrame(left_frame, text="处理进度", padx=8, pady=8)
        progress_frame.pack(fill="x", pady=5)

        self.three_d_progress_var = tk.DoubleVar()
        self.three_d_progress_bar = ttk.Progressbar(progress_frame, variable=self.three_d_progress_var, maximum=100)
        self.three_d_progress_bar.pack(fill="x", pady=5)

        self.three_d_progress_label = tk.Label(progress_frame, text="就绪", fg="green", font=("Arial", 9))
        self.three_d_progress_label.pack()

        # === 右侧区域：结果显示 ===
        right_frame = tk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # 结果显示区域
        result_frame = tk.LabelFrame(right_frame, text="处理结果", padx=8, pady=8)
        result_frame.pack(fill="both", expand=True)

        # 使用ScrolledText显示结果
        self.three_d_result_text_area = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Courier New", 9))
        self.three_d_result_text_area.pack(fill="both", expand=True)

        # 配置网格权重
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_rowconfigure(0, weight=1)

        # 检查pandas是否可用
        if not PANDAS_AVAILABLE:
            warning_text = "警告: pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl\n"
            self.three_d_result_text_area.insert(1.0, warning_text)
            self.three_d_start_button.config(state=tk.DISABLED)

    def browse_three_d_files(self):
        """浏览并选择多个Excel文件"""
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl")
            return

        filenames = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if filenames:
            for filename in filenames:
                if filename not in self.three_d_file_paths:
                    self.three_d_file_paths.append(filename)
                    # 添加到Treeview
                    index = len(self.three_d_file_paths)
                    self.three_d_file_tree.insert("", "end", values=(
                        index,
                        os.path.basename(filename)
                    ))

            # 更新统计信息
            self.update_three_d_file_count()

            # 自动设置输出文件夹为第一个文件所在目录
            if self.three_d_file_paths and (not self.three_d_output_dir.get() or self.three_d_output_dir.get() == os.path.expanduser("~")):
                input_dir = os.path.dirname(self.three_d_file_paths[0])
                self.three_d_output_dir.set(input_dir)

            self.status_bar.config(text=f"已添加 {len(filenames)} 个Excel文件")

    def browse_three_d_output_dir(self):
        """浏览并选择输出文件夹"""
        dirname = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self.three_d_output_dir.get()
        )
        if dirname:
            self.three_d_output_dir.set(dirname)
            self.status_bar.config(text=f"输出文件夹: {dirname}")

    def clear_three_d_file_list(self):
        """清空文件列表"""
        if self.three_d_file_paths:
            if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
                self.three_d_file_paths.clear()
                # 清空Treeview
                for item in self.three_d_file_tree.get_children():
                    self.three_d_file_tree.delete(item)
                self.update_three_d_file_count()
                self.status_bar.config(text="已清空文件列表")

    def remove_selected_three_d_files(self):
        """移除选中的文件"""
        selected_items = self.three_d_file_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要移除的文件")
            return

        # 确认移除
        if messagebox.askyesno("确认", f"确定要移除选中的 {len(selected_items)} 个文件吗？"):
            for item in selected_items:
                values = self.three_d_file_tree.item(item, "values")
                if values:
                    # 从文件路径列表中移除
                    filename = values[1]
                    # 需要找到完整路径
                    for file_path in self.three_d_file_paths[:]:
                        if os.path.basename(file_path) == filename:
                            self.three_d_file_paths.remove(file_path)
                            break
                # 从Treeview中移除
                self.three_d_file_tree.delete(item)

            # 重新排序
            self.reorder_three_d_file_list()
            self.update_three_d_file_count()
            self.status_bar.config(text=f"已移除 {len(selected_items)} 个文件")

    def reorder_three_d_file_list(self):
        """重新排序文件列表"""
        # 清空Treeview
        for item in self.three_d_file_tree.get_children():
            self.three_d_file_tree.delete(item)

        # 重新添加
        for i, filename in enumerate(self.three_d_file_paths, 1):
            self.three_d_file_tree.insert("", "end", values=(
                i,
                os.path.basename(filename)
            ))

    def update_three_d_file_count(self):
        """更新文件计数"""
        count = len(self.three_d_file_paths)
        self.three_d_file_count_label.config(text=f"已选择 {count} 个文件")

        # 根据文件数量启用/禁用处理按钮
        if count > 0 and PANDAS_AVAILABLE:
            self.three_d_start_button.config(state=tk.NORMAL)
        else:
            self.three_d_start_button.config(state=tk.DISABLED)

    def show_three_d_file_context_menu(self, event):
        """显示文件右键菜单"""
        # 选择右键点击的项目
        item = self.three_d_file_tree.identify_row(event.y)
        if item:
            self.three_d_file_tree.selection_set(item)

            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="打开文件", command=lambda: self.open_selected_three_d_file())
            menu.add_command(label="打开所在文件夹", command=lambda: self.open_three_d_file_folder())
            menu.add_separator()
            menu.add_command(label="从列表中移除", command=lambda: self.remove_selected_three_d_files())
            menu.tk_popup(event.x_root, event.y_root)

    def open_selected_three_d_file(self):
        """打开选中的文件"""
        selected_items = self.three_d_file_tree.selection()
        if selected_items:
            values = self.three_d_file_tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in self.three_d_file_paths:
                    if os.path.basename(file_path) == filename:
                        try:
                            if os.name == 'nt':
                                os.startfile(file_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', file_path] if os.sys.platform == 'darwin' else ['xdg-open', file_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件:\n{str(e)}")
                        break

    def open_three_d_file_folder(self):
        """打开文件所在文件夹"""
        selected_items = self.three_d_file_tree.selection()
        if selected_items:
            values = self.three_d_file_tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in self.three_d_file_paths:
                    if os.path.basename(file_path) == filename:
                        folder_path = os.path.dirname(file_path)
                        try:
                            if os.name == 'nt':
                                os.startfile(folder_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', folder_path] if os.sys.platform == 'darwin' else ['xdg-open', folder_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件夹:\n{str(e)}")
                        break

    def start_three_d_processing(self):
        """开始三次元数据处理"""
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl")
            return

        if not self.three_d_file_paths:
            messagebox.showerror("错误", "请先选择要处理的Excel文件！")
            return

        # 检查输出文件夹
        if not self.three_d_output_dir.get():
            messagebox.showerror("错误", "请选择输出文件夹！")
            return

        # 创建输出文件夹（如果需要）
        try:
            os.makedirs(self.three_d_output_dir.get(), exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹:\n{str(e)}")
            return

        # 禁用处理按钮，启用停止按钮
        self.toggle_three_d_buttons(processing=True)
        self.three_d_processing = True

        # 在新线程中处理
        thread = threading.Thread(target=self.process_three_d_files_thread, daemon=True)
        thread.start()

    def stop_three_d_processing(self):
        """停止处理"""
        self.three_d_processing = False
        self.status_bar.config(text="正在停止处理...")

    def toggle_three_d_buttons(self, processing):
        """切换按钮状态"""
        if processing:
            self.three_d_start_button.config(state=tk.DISABLED)
            self.three_d_stop_button.config(state=tk.NORMAL)
        else:
            self.three_d_start_button.config(state=tk.NORMAL if self.three_d_file_paths and PANDAS_AVAILABLE else tk.DISABLED)
            self.three_d_stop_button.config(state=tk.DISABLED)

    def process_three_d_files_thread(self):
        """处理三次元文件的线程函数"""
        try:
            # 清空结果区域
            self.three_d_result_text_area.delete(1.0, tk.END)

            # 显示开始信息
            self.show_three_d_result_header(len(self.three_d_file_paths))

            # 汇总所有文件的数据
            success, message, output_file = self.merge_three_d_files()

            if success:
                self.append_three_d_result(f"✓ {message}\n")
                self.three_d_progress_var.set(100)
                self.three_d_progress_label.config(text="汇总完成")
                self.status_bar.config(text=f"汇总完成: {os.path.basename(output_file)}")
            else:
                self.append_three_d_result(f"✗ {message}\n")
                self.three_d_progress_label.config(text="汇总失败")
                self.status_bar.config(text="汇总失败")

        except Exception as e:
            self.append_three_d_result(f"✗ 处理失败: {str(e)}\n")
        finally:
            # 处理完成
            self.three_d_processing = False
            self.toggle_three_d_buttons(processing=False)

    def show_three_d_result_header(self, total_files):
        """显示结果头部信息"""
        result_info = f"{'=' * 70}\n"
        result_info += f"三次元数据汇总开始 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        result_info += f"{'=' * 70}\n\n"
        result_info += f"文件总数: {total_files}\n"
        result_info += f"输出文件夹: {self.three_d_output_dir.get()}\n\n"
        result_info += f"{'=' * 70}\n"
        result_info += "处理结果:\n"
        result_info += f"{'=' * 70}\n\n"

        self.three_d_result_text_area.insert(1.0, result_info)

    def append_three_d_result(self, message):
        """在结果区域追加信息"""
        self.three_d_result_text_area.insert(tk.END, message)
        self.three_d_result_text_area.see(tk.END)
        self.root.update()

    def merge_three_d_files(self):
        """合并多个三次元Excel文件"""
        try:
            # 存储所有文件的数据
            all_data = {}  # {(尺寸, 轴): {nominal, upper_tol, lower_tol, measurements: []}}

            # 读取所有文件
            for file_idx, file_path in enumerate(self.three_d_file_paths, 1):
                if not self.three_d_processing:
                    return False, "处理已停止", None

                self.three_d_progress_var.set((file_idx / len(self.three_d_file_paths)) * 90)
                self.three_d_progress_label.config(text=f"正在读取文件 {file_idx}/{len(self.three_d_file_paths)}")
                self.status_bar.config(text=f"正在读取: {os.path.basename(file_path)}")
                self.root.update()

                # 读取Excel文件（尝试读取所有工作表）
                try:
                    # 先尝试读取第一个工作表
                    df = pd.read_excel(file_path, header=None, sheet_name=0)
                except Exception as e:
                    self.append_three_d_result(f"警告: {os.path.basename(file_path)} 读取失败: {str(e)}\n")
                    continue

                # 检查文件是否有数据
                if df.empty:
                    self.append_three_d_result(f"警告: {os.path.basename(file_path)} 文件为空\n")
                    continue

                # 查找数据开始行（跳过标题行）
                data_start_row = self.find_data_start_row(df)

                if data_start_row == -1:
                    self.append_three_d_result(f"警告: {os.path.basename(file_path)} 未找到数据行\n")
                    continue

                # 解析列索引
                col_indices = self.detect_column_indices(df, data_start_row)

                if not col_indices:
                    self.append_three_d_result(f"警告: {os.path.basename(file_path)} 无法识别列结构\n")
                    continue

                # 提取数据
                # 检查文件列数，确保能够访问所需的列（至少需要8列）
                max_cols = max([len(row) for row in df.values] + [len(df.columns)]) if len(df) > 0 else len(df.columns)
                if max_cols < 8:
                    self.append_three_d_result(f"警告: {os.path.basename(file_path)} 列数不足（需要至少8列，实际约{max_cols}列），跳过此文件\n")
                    continue
                
                for row_idx in range(data_start_row, len(df)):
                    row = df.iloc[row_idx]
                    
                    # 检查行是否有足够的列
                    if len(row) < 8:
                        continue
                    
                    # 只处理第1列（尺寸）、第2列（描述）和第4列（轴）都有数据的行
                    size = self.get_cell_value(row, col_indices['size'])         # 第1列：尺寸
                    description = self.get_cell_value(row, col_indices['description']) # 第2列：描述
                    axis = self.get_cell_value(row, col_indices['axis'])         # 第4列：轴

                    # 跳过无效行：第1列、第2列或第4列为空的行
                    if size is None or pd.isna(size) or str(size).strip() == '':
                        continue
                    if description is None or pd.isna(description) or str(description).strip() == '':
                        continue
                    if axis is None or pd.isna(axis) or str(axis).strip() == '':
                        continue

                    # 提取其他列的数据
                    nominal = self.get_cell_value(row, col_indices['nominal'])         # 第5列：标准值
                    upper_tol = self.get_cell_value(row, col_indices['upper_tol'])    # 第7列：上公差
                    lower_tol = self.get_cell_value(row, col_indices['lower_tol'])    # 第8列：下公差
                    measured = self.get_cell_value(row, col_indices['measured'])        # 第6列：实测值

                    # 创建键（尺寸+描述+轴）
                    key = (str(size).strip(), str(description).strip(), str(axis).strip())

                    # 如果是第一次遇到这个键，初始化数据
                    if key not in all_data:
                        all_data[key] = {
                            'description': description,
                            'nominal': nominal,
                            'upper_tol': upper_tol,
                            'lower_tol': lower_tol,
                            'measurements': []
                        }

                    # 添加实测值（第6列的数据）
                    if measured is not None and not pd.isna(measured):
                        all_data[key]['measurements'].append(measured)

            # 生成输出文件
            self.three_d_progress_label.config(text="正在生成输出文件...")
            self.root.update()

            output_file = self.generate_three_d_output(all_data)

            return True, f"汇总完成，共处理 {len(self.three_d_file_paths)} 个文件 -> {os.path.basename(output_file)}", output_file

        except Exception as e:
            return False, f"处理失败: {str(e)}", None

    def find_data_start_row(self, df):
        """查找数据开始行（跳过标题行）"""
        # 通常第一行是标题，第二行开始是数据
        # 但我们需要检查第一行是否真的是标题
        for i in range(min(10, len(df))):  # 增加搜索范围到10行
            row = df.iloc[i]
            # 检查是否包含"尺寸"、"轴"等关键词
            row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
            row_str_upper = row_str.upper()
            
            # 如果包含标题关键词，下一行是数据
            if any(keyword in row_str for keyword in ['尺寸', '轴', '标准值', '上公差', '下公差', '实测值', '特征', '描述']) or \
               any(keyword in row_str_upper for keyword in ['SIZE', 'AXIS', 'NOMINAL', 'TOL', 'MEAS', 'FEATURE', 'DESCRIPTION']):
                        # 检查下一行是否是数据行（包含数字或有效数据）
                        if i + 1 < len(df):
                            next_row = df.iloc[i + 1]
                            # 检查下一行是否包含有效数据（前几列有值）
                            if len(next_row) >= 4:
                                # 检查第1列、第2列、第4列是否有值（尺寸、描述、轴）
                                if pd.notna(next_row.iloc[0]) and pd.notna(next_row.iloc[1]) and pd.notna(next_row.iloc[3]):
                                    return i + 1  # 下一行开始是数据
                            elif len(next_row) >= 2:
                                # 如果列数不够，至少检查前两列
                                if pd.notna(next_row.iloc[0]) and pd.notna(next_row.iloc[1]):
                                    return i + 1  # 下一行开始是数据
                        return i + 1  # 即使下一行可能为空，也返回下一行
        
        # 如果没找到标题行，尝试从第一行开始查找数据
        # 检查第一行是否是数据行（包含数字）
        if len(df) > 0:
            first_row = df.iloc[0]
            if len(first_row) >= 4:
                # 检查第1列、第2列、第4列是否有值（尺寸、描述、轴）
                if pd.notna(first_row.iloc[0]) and pd.notna(first_row.iloc[1]) and pd.notna(first_row.iloc[3]):
                    return 0  # 第一行就是数据
        
        # 默认从第二行开始（索引1），如果只有一行则从第一行开始
        return 1 if len(df) > 1 else 0

    def detect_column_indices(self, df, header_row):
        """检测列索引，使用固定的列索引映射"""
        # 根据用户要求，使用固定的列索引：
        # 汇总文件第1列（尺寸）：源文件第1列（索引0）
        # 汇总文件第2列（描述）：源文件第2列（索引1）
        # 汇总文件第3列（特征）：源文件第3列（索引2）
        # 汇总文件第4列（轴）：源文件第4列（索引3）
        # 汇总文件第5列（标准值NOMINAL）：源文件第5列（索引4）
        # 汇总文件第6列（上公差+TOL）：源文件第7列（索引6）
        # 汇总文件第7列（下公差-TOL）：源文件第8列（索引7）
        # 汇总文件第8列开始（实测值MEAS）：源文件第6列（索引5）
        
        col_indices = {
            'size': 0,        # 第1列：尺寸
            'description': 1,  # 第2列：描述
            'feature': 2,     # 第3列：特征
            'axis': 3,        # 第4列：轴
            'nominal': 4,     # 第5列：标准值
            'upper_tol': 6,   # 第7列：上公差
            'lower_tol': 7,   # 第8列：下公差
            'measured': 5     # 第6列：实测值
        }
        
        return col_indices

    def get_cell_value(self, row, col_idx):
        """获取单元格值"""
        if col_idx >= len(row):
            return None
        value = row.iloc[col_idx]
        return value if pd.notna(value) else None

    def generate_three_d_output(self, all_data):
        """生成汇总后的Excel文件"""
        # 准备输出数据
        output_rows = []

        # 标题行（根据用户要求修改）
        header = ['尺寸', '描述', '轴', 'NOMINAL', '+TOL', '-TOL']
        # 计算最大实测值数量
        max_measurements = max(len(data['measurements']) for data in all_data.values()) if all_data else 0
        # 添加实测值列标题（从第7列开始）
        for i in range(max_measurements):
            header.append('MEAS')

        output_rows.append(header)

        # 数据行（按尺寸、描述和轴排序）
        def get_sort_key(key):
            """生成排序键，与OGP数据排序方式一致"""
            size, description, axis = key
            size_str = str(size).strip()
            
            # 初始化排序键
            sort_key = None
            
            try:
                # 检查是否以数字开头
                if size_str and size_str[0].isdigit():
                    # 数字开头的标签
                    if '*' in size_str:
                        # 处理带星号的标签，如 "15*1", "15*2"
                        num_parts = size_str.split('*')
                        main_num = int(num_parts[0])
                        sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0
                        sort_key = (0, '', main_num, sub_num, description, axis)
                    else:
                        # 处理纯数字标签，如 "1", "2", "11"
                        main_num = int(size_str)
                        sort_key = (0, '', main_num, 0, description, axis)
                else:
                    # 字母开头的标签
                    # 提取字母部分和数字部分
                    alpha_part = ''
                    num_part = ''
                    star_part = ''
                    
                    # 找到第一个数字的位置
                    for i, char in enumerate(size_str):
                        if char.isdigit():
                            alpha_part = size_str[:i]
                            remaining = size_str[i:]
                            # 处理可能的星号
                            if '*' in remaining:
                                num_star_parts = remaining.split('*')
                                num_part = num_star_parts[0]
                                star_part = num_star_parts[1] if len(num_star_parts) > 1 else '0'
                            else:
                                num_part = remaining
                                star_part = '0'
                            break
                    
                    # 转换数字部分为整数
                    try:
                        num_val = int(num_part)
                        star_val = int(star_part)
                    except ValueError:
                        num_val = 0
                        star_val = 0
                    
                    # 字母开头的排序键：(1, 字母部分, 数字部分, 星号部分, 描述, 轴)
                    sort_key = (1, alpha_part, num_val, star_val, description, axis)
            except (ValueError, IndexError):
                # 如果解析失败，放到最后
                sort_key = (2, '', 0, 0, description, axis)
            
            return sort_key
        
        sorted_keys = sorted(all_data.keys(), key=get_sort_key)

        for key in sorted_keys:
            size, description, axis = key
            data = all_data[key]

            row = [
                size,  # 第1列：尺寸（源文件第1列）
                description,  # 第2列：描述（源文件第2列）
                axis,  # 第3列：轴（源文件第4列）
                data['nominal'] if pd.notna(data['nominal']) else '',  # 第4列：NOMINAL（源文件第5列）
                data['upper_tol'] if pd.notna(data['upper_tol']) else '',  # 第5列：+TOL（源文件第7列）
                data['lower_tol'] if pd.notna(data['lower_tol']) else ''  # 第6列：-TOL（源文件第8列）
            ]

            # 添加所有实测值（第8列开始，来源于源文件第6列）
            for meas in data['measurements']:
                row.append(meas if pd.notna(meas) else '')

            # 如果实测值数量少于最大值，用空值填充
            while len(row) < len(header):
                row.append('')

            output_rows.append(row)

        # 创建DataFrame
        output_df = pd.DataFrame(output_rows)

        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"三次元汇总_{timestamp}.xlsx"
        output_file = os.path.join(self.three_d_output_dir.get(), output_filename)

        # 保存为Excel文件
        output_df.to_excel(output_file, index=False, header=False, engine='openpyxl')

        return output_file


def detect_encoding(file_path):
    """检测文件编码"""
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(4096)
            result = chardet.detect(raw_data)
            # result = chardet.detect(f.read())

            encoding = result['encoding']
            confidence = result['confidence']

            if not encoding or confidence < 0.7:
                encoding = 'utf-8'
            print("****************", encoding, confidence)
            return encoding
    except Exception as e:
        print(f"编码检测失败: {e}")
        return 'utf-8'


def main():
    root = tk.Tk()

    # 配置进度条样式
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("red.Horizontal.TProgressbar",
                    troughcolor='white',
                    background='red',
                    bordercolor='gray',
                    lightcolor='red',
                    darkcolor='red')

    app = DataSorterApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        print(f"错误: 缺少必要的库 - {e}")
        print("请运行: pip install chardet")
        input("按Enter键退出...")