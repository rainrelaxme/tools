#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
@Project : tools
@File    : ogp2.py
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

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 创建主框架，使用grid布局
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # === 左侧区域：文件选择和参数设置 ===
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # 处理模式选择
        mode_frame = tk.LabelFrame(left_frame, text="处理模式", padx=8, pady=8)
        mode_frame.pack(fill="x", pady=(0, 5))

        self.process_mode = tk.StringVar(value="summary")  # 默认排序模式

        tk.Radiobutton(mode_frame, text="仅排序", variable=self.process_mode,
                       value="sort", font=("Arial", 9)).pack(anchor="w", pady=2)
        tk.Radiobutton(mode_frame, text="排序并汇总", variable=self.process_mode,
                       value="summary", font=("Arial", 9)).pack(anchor="w", pady=2)

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

        # 汇总模式说明
        # summary_info_frame = tk.Frame(right_frame)
        # summary_info_frame.pack(fill="x", pady=(5, 0))
        #
        # tk.Label(summary_info_frame, text="汇总模式说明:", font=("Arial", 9, "bold")).pack(anchor="w")
        # info_text = tk.Text(summary_info_frame, height=3, wrap=tk.WORD, font=("Arial", 8), bg="#f0f0f0")
        # info_text.pack(fill="x", pady=2)
        # info_text.insert(1.0,
        #                  "1. 按区块排序数据\n2. 调整列顺序并合并实测值\n3. 只保留第一个区块，其他区块的实测值追加到第一区块")
        # info_text.config(state=tk.DISABLED)

        # 状态栏
        self.status_bar = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

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
                content = f.read()

            # 根据处理模式选择处理方法
            if self.process_mode.get() == "summary":
                # 汇总模式
                result, block_count = self.summarize_data(content)
                mode_text = "汇总"
            else:
                # 仅排序模式
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

    def show_result_header(self, total_files):
        """显示结果头部信息"""
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"

        result_info = f"{'=' * 70}\n"
        result_info += f"批量{mode_text}开始 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        result_info += f"{'=' * 70}\n\n"
        result_info += f"文件总数: {total_files}\n"
        result_info += f"输出文件夹: {self.output_dir.get()}\n"
        result_info += f"表头行数: {self.header_lines.get()}\n"
        result_info += f"处理模式: {mode_text}\n\n"
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
        """仅排序模式：处理文件内容，按区块排序第一列数据"""
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块
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

    def summarize_data(self, content):
        """
        汇总模式：
        1. 先排序数据
        2. 调整列顺序并合并实测值
        3. 只保留第一个区块
        """
        lines = content.strip().split('\n')
        blocks = []
        current_block = []
        data_line_pattern = re.compile(r'^\s*(\d+[\*\d]*)\s+')

        # 识别区块
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
                first_col = match.group(1)
                if '*' in first_col:
                    try:
                        parts = first_col.split('*')
                        key = (int(parts[0]), int(parts[1]))
                    except ValueError:
                        key = (float('inf'), i)
                else:
                    try:
                        key = (int(first_col), 0)
                    except ValueError:
                        key = (float('inf'), i)
                data_with_keys.append((key, line))
            else:
                data_with_keys.append(((float('inf'), i), line))

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
        # else:
        #     # 尝试用制表符分割
        #     columns = line.strip().split('\t')
        #     if len(columns) >= 6:
        #         return columns

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

                if '*' in main_part:
                    # 处理带星号的标签，如 "4*1", "20*20"
                    num_parts = main_part.split('*')
                    main_num = int(num_parts[0])
                    sub_num = int(num_parts[1]) if len(num_parts) > 1 else 0

                    # 特殊处理：对于 20*20 X，我们希望它排在 20*3 X 之后
                    # 所以需要将 sub_num 作为主要排序因素之一
                    sort_key = (main_num, sub_num, 1 if suffix else 0, suffix, label_stripped)
                else:
                    # 处理普通数字标签，如 "22", "6 X"
                    main_num = int(main_part)
                    sort_key = (main_num, 0, 1 if suffix else 0, suffix, label_stripped)

            except (ValueError, IndexError):
                # 如果解析失败，放到最后
                sort_key = (float('inf'), 0, 0, '', label_stripped)

            data_list.append((sort_key, label, block_data))

        # 按标签排序
        data_list.sort(key=lambda x: x[0])

        return data_list

    # def sort_extracted_data(self, data_dict):
    #     """对提取的数据进行排序"""
    #     # 将数据转换为列表以便排序
    #     data_list = []
    #     for label, block_data in data_dict.items():
    #         # 解析标签以便排序
    #         if '*' in label:
    #             try:
    #                 parts = label.split('*')
    #                 sort_key = (int(parts[0]), int(parts[1]))
    #             except ValueError:
    #                 sort_key = (float('inf'), 0)
    #         else:
    #             try:
    #                 sort_key = (int(label), 0)
    #             except ValueError:
    #                 sort_key = (float('inf'), 0)
    #
    #         data_list.append((sort_key, label, block_data))
    #
    #     # 按标签排序
    #     data_list.sort(key=lambda x: x[0])
    #
    #     return data_list


    # def sort_extracted_data(self, data_dict):
    #     """对提取的数据进行排序"""
    #     data_list = []
    #
    #     for i, (label, block_data) in enumerate(data_dict.items()):
    #         # 解析排序键（与sort_data_lines完全相同的逻辑）
    #         if '*' in label:
    #             try:
    #                 parts = label.split('*')
    #                 sort_key = (int(parts[0]), int(parts[1]))
    #             except ValueError:
    #                 sort_key = (float('inf'), i)
    #         else:
    #             try:
    #                 sort_key = (int(label), 0)
    #             except ValueError:
    #                 sort_key = (float('inf'), i)
    #
    #         data_list.append((sort_key, label, block_data))
    #
    #     # 排序
    #     data_list.sort(key=lambda x: x[0])
    #
    #     # 返回排序后的字典
    #     return {label: block_data for _, label, block_data in data_list}

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
        # 排序标题行
        title = result[-1].split('\t')
        label = title[0]
        dim_type = title[1]
        nominal = title[2]
        measured = title[3]
        upper_tol = title[4]
        lower_tol = title[5]
        other = title[6:]
        formatted_title = f"{label}\t{dim_type}\t{nominal}\t{upper_tol}\t{lower_tol}\t{measured}"
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
                other_cols = base_columns[6:]  # 第6,7,8列

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


def detect_encoding(file_path):
    """检测文件编码"""
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(4096)
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            confidence = result['confidence']

            if not encoding or confidence < 0.7:
                encoding = 'utf-8'
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