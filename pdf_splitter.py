#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF拆分工具 v1.2
支持按章节（书签）拆分、自定义页码范围拆分和按大小均匀拆分
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import shutil
import threading
import re
import subprocess

# 优先使用 ttkbootstrap（现代 UI），否则回退到标准 ttk
try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    MODERN_UI = True
except ImportError:
    from tkinter import ttk
    MODERN_UI = False

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def sanitize_filename(name):
    """移除文件名中的非法字符"""
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.strip(' .')
    return name if name else '未命名'


def _get_platform_fonts():
    """根据操作系统返回合适的字体"""
    if sys.platform == 'darwin':
        return 'PingFang SC', 'Menlo'
    elif sys.platform == 'win32':
        return 'Microsoft YaHei UI', 'Consolas'
    else:  # Linux
        return 'Noto Sans CJK SC', 'Monospace'


UI_FONT, MONO_FONT = _get_platform_fonts()


def _bs(style_str):
    """返回 bootstyle 参数（仅 ttkbootstrap 可用时生效）"""
    return {'bootstyle': style_str} if MODERN_UI else {}


class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 拆分工具 v1.2")
        self.root.minsize(800, 700)

        # ---------- 状态变量 ----------
        self.pdf_path = None
        self.pdf_doc = None
        self.total_pages = 0
        self.chapters = []
        self.split_mode = tk.StringVar(value="chapter")
        self.chapter_vars = []          # [(BooleanVar, chapter_dict), ...]
        self.custom_rows = []           # [(frame, start_entry, end_entry, name_entry), ...]
        self.size_split_tasks = []      # 按大小拆分的任务列表
        self.max_size_mb = tk.StringVar(value="200")
        self.level_var = tk.StringVar(value="1")
        self.output_dir = None
        self.is_running = False

        # ---------- 构建界面 ----------
        self._build_ui()
        self._center_window(860, 740)

    # ================================================================
    #  界面构建
    # ================================================================

    def _center_window(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def _build_ui(self):
        # 可滚动的主容器
        outer = ttk.Frame(self.root)
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        self.main_frame = ttk.Frame(canvas)

        self.main_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.main_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮支持（跨平台）
        if sys.platform == 'darwin':
            canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(-e.delta, 'units'))
        elif sys.platform == 'linux':
            canvas.bind_all('<Button-4>', lambda e: canvas.yview_scroll(-1, 'units'))
            canvas.bind_all('<Button-5>', lambda e: canvas.yview_scroll(1, 'units'))
        else:  # Windows
            canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units'))

        # 内容区（用 tk.Frame 包裹加内边距，避免 ttk padding 兼容问题）
        content = tk.Frame(self.main_frame)
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self._content = content

        # 标题
        tk.Label(content, text="PDF 拆分工具",
                 font=(UI_FONT, 20, 'bold'), fg='#0d6efd').pack(pady=(0, 3))
        tk.Label(content, text="轻松将一本PDF拆分为多个小文件，原文件不会被修改",
                 font=(UI_FONT, 10), fg='#6c757d').pack(pady=(0, 18))

        self._build_step1(content)
        self._build_step2(content)
        self._build_step3(content)
        self._build_log(content)

    # ---------- 第一步：选择文件 ----------

    def _build_step1(self, parent):
        frame = tk.LabelFrame(parent, text="  第一步：选择要拆分的PDF文件  ",
                              font=(UI_FONT, 10), padx=12, pady=12)
        frame.pack(fill=tk.X, pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill=tk.X)

        ttk.Button(row, text="  选择PDF文件 ...  ", command=self._select_file,
                   **_bs('primary-outline')).pack(side=tk.LEFT)

        self.file_info_label = tk.Label(row, text="  尚未选择文件",
                                        font=(UI_FONT, 10), fg='gray')
        self.file_info_label.pack(side=tk.LEFT, padx=(15, 0))

    # ---------- 第二步：选择拆分方式 ----------

    def _build_step2(self, parent):
        frame = tk.LabelFrame(parent, text="  第二步：选择拆分方式  ",
                              font=(UI_FONT, 10), padx=12, pady=12)
        frame.pack(fill=tk.X, pady=(0, 10))

        # 模式选择
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Radiobutton(mode_frame, text="按书签/章节拆分（自动识别PDF中的目录结构）",
                        variable=self.split_mode, value="chapter",
                        command=self._on_mode_change).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="自定义页码范围拆分（手动指定每一部分的起止页码）",
                        variable=self.split_mode, value="custom",
                        command=self._on_mode_change).pack(anchor=tk.W, pady=(4, 0))
        ttk.Radiobutton(mode_frame, text="按大小均匀拆分（无书签时自动建议，每份不超过指定大小）",
                        variable=self.split_mode, value="size",
                        command=self._on_mode_change).pack(anchor=tk.W, pady=(4, 0))

        # 内容面板容器
        self.panel_container = ttk.Frame(frame)
        self.panel_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        # 章节面板
        self.chapter_panel = ttk.Frame(self.panel_container)
        self._build_chapter_panel()

        # 自定义面板
        self.custom_panel = ttk.Frame(self.panel_container)
        self._build_custom_panel()

        # 按大小拆分面板
        self.size_panel = ttk.Frame(self.panel_container)
        self._build_size_panel()

        # 默认显示章节面板
        self.chapter_panel.pack(fill=tk.BOTH, expand=True)

    def _build_chapter_panel(self):
        # 层级选择
        top_row = ttk.Frame(self.chapter_panel)
        top_row.pack(fill=tk.X, pady=(0, 6))

        tk.Label(top_row, text="拆分层级：", font=(UI_FONT, 10)).pack(side=tk.LEFT)
        level_combo = ttk.Combobox(top_row, textvariable=self.level_var, width=25, state='readonly',
                                   values=["1", "2", "3"])
        level_combo.pack(side=tk.LEFT, padx=(4, 0))
        level_combo.bind('<<ComboboxSelected>>', lambda e: self._refresh_chapters())

        tk.Label(top_row, text="（1=按章拆分，2=按节拆分，3=更细）",
                 font=(UI_FONT, 9), fg='#6c757d').pack(side=tk.LEFT, padx=(8, 0))

        # 按钮行
        btn_row = ttk.Frame(self.chapter_panel)
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_row, text="全选", command=lambda: self._toggle_all(True),
                   **_bs('secondary-outline')).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="取消全选", command=lambda: self._toggle_all(False),
                   **_bs('secondary-outline')).pack(side=tk.LEFT)

        # 章节列表（带滚动条）
        list_frame = ttk.Frame(self.chapter_panel)
        list_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_frame, height=200, highlightthickness=1, highlightbackground='#cccccc')
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.chapter_inner = ttk.Frame(canvas)

        self.chapter_inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.chapter_inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.chapter_canvas = canvas

        # 提示
        self.chapter_hint = tk.Label(self.chapter_panel, text="请先选择一个PDF文件",
                                     font=(UI_FONT, 10), fg='gray')
        self.chapter_hint.pack(pady=8)

    def _build_custom_panel(self):
        tk.Label(self.custom_panel,
                 text="在下方添加要拆分的页码范围，\"名称\"可不填（将自动编号）。\n"
                      "页码从1开始，例如：第1页到第10页，就填  起始页=1  结束页=10",
                 font=(UI_FONT, 9), fg='#6c757d', justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 8))

        # 表头
        header = ttk.Frame(self.custom_panel)
        header.pack(fill=tk.X)
        tk.Label(header, text="起始页", width=10, anchor=tk.CENTER,
                 font=(UI_FONT, 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(header, text="结束页", width=10, anchor=tk.CENTER,
                 font=(UI_FONT, 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(header, text="名称（可不填）", width=25, anchor=tk.CENTER,
                 font=(UI_FONT, 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))

        # 行容器
        self.custom_rows_frame = ttk.Frame(self.custom_panel)
        self.custom_rows_frame.pack(fill=tk.X, pady=(4, 0))

        # 添加按钮
        ttk.Button(self.custom_panel, text="+ 添加一行", command=self._add_range_row,
                   **_bs('info-outline')).pack(anchor=tk.W, pady=(6, 0))

        # 默认添加一行
        self._add_range_row()

    def _build_size_panel(self):
        """构建按大小拆分面板"""
        # 说明文字
        tk.Label(self.size_panel,
                 text="自动根据文件大小计算拆分方案，确保每份不超过指定大小。\n"
                      "适合没有书签的大PDF文件，省去手动计算页码的麻烦。",
                 font=(UI_FONT, 9), fg='#6c757d', justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 10))

        # 设置行：最大文件大小
        setting_row = ttk.Frame(self.size_panel)
        setting_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(setting_row, text="每份最大大小：", font=(UI_FONT, 10)).pack(side=tk.LEFT)
        size_entry = ttk.Entry(setting_row, textvariable=self.max_size_mb, width=8, justify=tk.CENTER)
        size_entry.pack(side=tk.LEFT, padx=(4, 4))
        tk.Label(setting_row, text="MB", font=(UI_FONT, 10)).pack(side=tk.LEFT)

        ttk.Button(setting_row, text="重新计算建议", command=self._calc_size_split,
                   **_bs('warning-outline')).pack(side=tk.LEFT, padx=(20, 0))

        # 文件信息区域
        self.size_info_frame = ttk.Frame(self.size_panel)
        self.size_info_frame.pack(fill=tk.X, pady=(0, 8))

        self.size_info_label = tk.Label(self.size_info_frame, text="请先选择一个PDF文件",
                                        font=(UI_FONT, 10), fg='gray')
        self.size_info_label.pack(anchor=tk.W)

        # 建议方案展示区域（带滚动条）
        list_frame = ttk.Frame(self.size_panel)
        list_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_frame, height=180, highlightthickness=1, highlightbackground='#cccccc')
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.size_inner = ttk.Frame(canvas)

        self.size_inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.size_inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.size_canvas = canvas

    # ---------- 第三步：开始拆分 ----------

    def _build_step3(self, parent):
        frame = tk.LabelFrame(parent, text="  第三步：开始拆分  ",
                              font=(UI_FONT, 10), padx=12, pady=12)
        frame.pack(fill=tk.X, pady=(0, 10))

        # 输出路径
        out_row = ttk.Frame(frame)
        out_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(out_row, text="输出位置：", font=(UI_FONT, 10)).pack(side=tk.LEFT)
        self.output_label = tk.Label(out_row, text="（与源文件同一目录，自动创建文件夹）",
                                     font=(UI_FONT, 10), fg='gray')
        self.output_label.pack(side=tk.LEFT, padx=(4, 0))

        # 操作按钮
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)

        self.start_btn = ttk.Button(btn_row, text="  开始拆分  ", command=self._start_split,
                                    **_bs('success'))
        self.start_btn.pack(side=tk.LEFT)

        self.open_folder_btn = ttk.Button(btn_row, text="  打开输出文件夹  ", command=self._open_output_folder,
                                          **_bs('info-outline'))
        self.open_folder_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.open_folder_btn.pack_forget()  # 初始隐藏

        # 进度条
        self.progress = ttk.Progressbar(frame, mode='determinate', length=400,
                                        **_bs('success-striped'))
        self.progress.pack(fill=tk.X, pady=(10, 4))

        self.status_label = tk.Label(frame, text="就绪", font=(UI_FONT, 10), fg='#555555')
        self.status_label.pack(anchor=tk.W)

    # ---------- 日志区域 ----------

    def _build_log(self, parent):
        frame = tk.LabelFrame(parent, text="  运行日志  ",
                              font=(UI_FONT, 10), padx=8, pady=8)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.log_text = tk.Text(frame, height=8, wrap=tk.WORD, font=(MONO_FONT, 9),
                                state=tk.DISABLED, bg='#fafafa')
        log_sb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_sb.pack(side=tk.RIGHT, fill=tk.Y)

    # ================================================================
    #  文件选择与加载
    # ================================================================

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择要拆分的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if not path:
            return

        self.pdf_path = path
        self._log(f"已选择文件：{path}")
        self._load_pdf()

    def _load_pdf(self):
        if not fitz:
            messagebox.showerror("缺少依赖",
                                 "未找到 PyMuPDF 库。\n\n"
                                 "请在命令行运行以下命令安装：\n"
                                 "pip install PyMuPDF")
            return

        try:
            if self.pdf_doc:
                self.pdf_doc.close()
            self.pdf_doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.pdf_doc)

            fname = os.path.basename(self.pdf_path)
            self.file_info_label.configure(
                text=f"  {fname}   （共 {self.total_pages} 页）",
                fg='black'
            )

            self.output_dir = os.path.join(
                os.path.dirname(self.pdf_path),
                sanitize_filename(os.path.splitext(fname)[0]) + "_拆分结果"
            )
            self.output_label.configure(text=self.output_dir, fg='black')

            self._refresh_chapters()
            self._log(f"文件加载成功，共 {self.total_pages} 页")

            # 如果没有书签，自动切换到按大小拆分模式
            file_size = os.path.getsize(self.pdf_path)
            file_size_mb = file_size / (1024 * 1024)
            self._log(f"文件大小：{file_size_mb:.1f} MB")

            toc = self.pdf_doc.get_toc()
            if not toc:
                self._log("未检测到书签/目录，已自动切换到\"按大小均匀拆分\"模式")
                self.split_mode.set("size")
                self._on_mode_change()

        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开PDF文件：\n{e}")
            self._log(f"打开文件失败：{e}")

    # ================================================================
    #  章节检测
    # ================================================================

    def _get_chapters(self, max_level=1):
        """从PDF书签中提取章节信息"""
        toc = self.pdf_doc.get_toc()
        if not toc:
            return []

        # 按层级过滤
        filtered = []
        for level, title, page in toc:
            if level <= max_level:
                filtered.append((title.strip(), max(1, page)))

        if not filtered:
            return []

        chapters = []
        for i, (title, start_page) in enumerate(filtered):
            if i + 1 < len(filtered):
                end_page = filtered[i + 1][1] - 1
            else:
                end_page = self.total_pages

            if end_page < start_page:
                end_page = start_page

            chapters.append({
                'title': title,
                'start': start_page,
                'end': end_page,
            })

        return chapters

    def _refresh_chapters(self):
        """刷新章节列表显示"""
        # 清空旧内容
        for widget in self.chapter_inner.winfo_children():
            widget.destroy()
        self.chapter_vars.clear()

        if not self.pdf_doc:
            self.chapter_hint.configure(text="请先选择一个PDF文件")
            return

        max_level = int(self.level_var.get())
        self.chapters = self._get_chapters(max_level)

        if not self.chapters:
            self.chapter_hint.configure(
                text="此PDF没有书签/目录信息，建议使用\"按大小均匀拆分\"或\"自定义页码范围拆分\""
            )
            return

        self.chapter_hint.configure(text=f"共检测到 {len(self.chapters)} 个章节，勾选要拆分的章节：")

        for i, ch in enumerate(self.chapters):
            var = tk.BooleanVar(value=True)
            row = ttk.Frame(self.chapter_inner)
            row.pack(fill=tk.X, padx=4, pady=1)

            cb = ttk.Checkbutton(row, variable=var)
            cb.pack(side=tk.LEFT)

            text = f"{ch['title']}    （第 {ch['start']} - {ch['end']} 页）"
            tk.Label(row, text=text, font=(UI_FONT, 10)).pack(side=tk.LEFT, padx=(4, 0))

            self.chapter_vars.append((var, ch))

    def _toggle_all(self, state):
        for var, _ in self.chapter_vars:
            var.set(state)

    # ================================================================
    #  按大小拆分计算
    # ================================================================

    def _calc_size_split(self):
        """根据文件大小和页数计算拆分建议"""
        # 清空旧内容
        for widget in self.size_inner.winfo_children():
            widget.destroy()
        self.size_split_tasks.clear()

        if not self.pdf_doc or not self.pdf_path:
            self.size_info_label.configure(text="请先选择一个PDF文件", fg='gray')
            return

        # 获取最大大小限制
        try:
            max_mb = float(self.max_size_mb.get().strip())
            if max_mb <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("输入错误", "请输入一个有效的正数作为最大文件大小（MB）！")
            return

        max_bytes = max_mb * 1024 * 1024
        file_size = os.path.getsize(self.pdf_path)
        file_size_mb = file_size / (1024 * 1024)

        # 如果文件本身就不超过限制
        if file_size <= max_bytes:
            self.size_info_label.configure(
                text=f"文件大小：{file_size_mb:.1f} MB，未超过 {max_mb:.0f} MB，无需拆分。",
                fg='green'
            )
            return

        avg_page_size = file_size / self.total_pages
        pages_per_part = max(1, int(max_bytes / avg_page_size))
        num_parts = -(-self.total_pages // pages_per_part)  # 向上取整

        tasks = []
        for i in range(num_parts):
            start = i * pages_per_part + 1
            end = min((i + 1) * pages_per_part, self.total_pages)
            est_size_mb = (end - start + 1) * avg_page_size / (1024 * 1024)
            name = f"第{i+1}部分_第{start}-{end}页"
            tasks.append((start, end, name, est_size_mb))

        self.size_split_tasks = [(s, e, n) for s, e, n, _ in tasks]

        # 显示信息
        self.size_info_label.configure(
            text=f"文件大小：{file_size_mb:.1f} MB  |  共 {self.total_pages} 页  |  "
                 f"平均每页 {avg_page_size/1024:.1f} KB\n"
                 f"建议拆分为 {num_parts} 份，每份约 {pages_per_part} 页（预估不超过 {max_mb:.0f} MB）",
            fg='#333333'
        )

        # 展示拆分方案
        for i, (start, end, name, est_mb) in enumerate(tasks):
            row = ttk.Frame(self.size_inner)
            row.pack(fill=tk.X, padx=4, pady=2)

            page_count = end - start + 1
            text = (f"第 {i+1} 份：第 {start} - {end} 页  "
                    f"（{page_count} 页，预估 {est_mb:.1f} MB）")

            tk.Label(row, text=text, font=(UI_FONT, 9)).pack(side=tk.LEFT, padx=(4, 0))

    # ================================================================
    #  自定义范围行管理
    # ================================================================

    def _add_range_row(self):
        row_frame = ttk.Frame(self.custom_rows_frame)
        row_frame.pack(fill=tk.X, pady=2)

        start_entry = ttk.Entry(row_frame, width=10, justify=tk.CENTER)
        start_entry.pack(side=tk.LEFT, padx=(0, 4))

        end_entry = ttk.Entry(row_frame, width=10, justify=tk.CENTER)
        end_entry.pack(side=tk.LEFT, padx=(0, 4))

        name_entry = ttk.Entry(row_frame, width=28)
        name_entry.pack(side=tk.LEFT, padx=(0, 4))

        remove_btn = ttk.Button(row_frame, text="删除", width=5,
                                command=lambda: self._remove_range_row(row_frame),
                                **_bs('danger-outline'))
        remove_btn.pack(side=tk.LEFT)

        self.custom_rows.append((row_frame, start_entry, end_entry, name_entry))

    def _remove_range_row(self, row_frame):
        self.custom_rows = [(f, s, e, n) for f, s, e, n in self.custom_rows if f != row_frame]
        row_frame.destroy()

    # ================================================================
    #  模式切换
    # ================================================================

    def _on_mode_change(self):
        self.chapter_panel.pack_forget()
        self.custom_panel.pack_forget()
        self.size_panel.pack_forget()

        if self.split_mode.get() == "chapter":
            self.chapter_panel.pack(fill=tk.BOTH, expand=True)
        elif self.split_mode.get() == "custom":
            self.custom_panel.pack(fill=tk.BOTH, expand=True)
        else:  # size
            self.size_panel.pack(fill=tk.BOTH, expand=True)
            self._calc_size_split()

    # ================================================================
    #  拆分逻辑
    # ================================================================

    def _start_split(self):
        if self.is_running:
            return

        if not self.pdf_path or not self.pdf_doc:
            messagebox.showwarning("提示", "请先选择一个PDF文件！")
            return

        # 收集拆分任务
        tasks = self._collect_tasks()
        if tasks is None:
            return

        if not tasks:
            messagebox.showwarning("提示", "没有要拆分的内容，请至少选择一个章节或添加一个页码范围！")
            return

        self.is_running = True
        self.start_btn.configure(state=tk.DISABLED)
        self.open_folder_btn.pack_forget()
        self.progress['value'] = 0

        threading.Thread(target=self._do_split, args=(tasks,), daemon=True).start()

    def _collect_tasks(self):
        """收集拆分任务，返回 [(start, end, name), ...] 或 None（出错）"""
        tasks = []

        if self.split_mode.get() == "chapter":
            selected = [(ch['start'], ch['end'], ch['title'])
                        for var, ch in self.chapter_vars if var.get()]
            if not selected:
                messagebox.showwarning("提示", "请至少勾选一个章节！")
                return None
            tasks = selected

        elif self.split_mode.get() == "size":
            if not self.size_split_tasks:
                messagebox.showwarning("提示", "没有拆分方案，请先点击\"重新计算建议\"！")
                return None
            tasks = self.size_split_tasks

        else:  # custom
            for i, (_, start_e, end_e, name_e) in enumerate(self.custom_rows):
                s_text = start_e.get().strip()
                e_text = end_e.get().strip()
                n_text = name_e.get().strip()

                if not s_text and not e_text:
                    continue  # 跳过空行

                if not s_text or not e_text:
                    messagebox.showwarning("输入错误", f"第 {i+1} 行的起始页和结束页都必须填写！")
                    return None

                try:
                    s = int(s_text)
                    e = int(e_text)
                except ValueError:
                    messagebox.showwarning("输入错误", f"第 {i+1} 行的页码必须是数字！")
                    return None

                if s < 1 or e < 1:
                    messagebox.showwarning("输入错误", f"第 {i+1} 行的页码必须大于0！")
                    return None

                if s > e:
                    messagebox.showwarning("输入错误", f"第 {i+1} 行的起始页（{s}）不能大于结束页（{e}）！")
                    return None

                if e > self.total_pages:
                    messagebox.showwarning("输入错误",
                                           f"第 {i+1} 行的结束页（{e}）超出了PDF总页数（{self.total_pages}）！")
                    return None

                name = n_text if n_text else f"第{i+1}部分_第{s}-{e}页"
                tasks.append((s, e, name))

            if not tasks:
                messagebox.showwarning("提示", "请至少填写一行有效的页码范围！")
                return None

        return tasks

    def _do_split(self, tasks):
        """在后台线程中执行拆分"""
        try:
            output_dir = self.output_dir
            backup_dir = os.path.join(output_dir, "备份")

            # 创建输出目录
            os.makedirs(backup_dir, exist_ok=True)
            self._log(f"输出目录：{output_dir}")

            # 备份原文件
            backup_path = os.path.join(backup_dir, os.path.basename(self.pdf_path))
            if not os.path.exists(backup_path):
                shutil.copy2(self.pdf_path, backup_path)
                self._log(f"已备份原文件到：备份/{os.path.basename(self.pdf_path)}")
            else:
                self._log("备份文件已存在，跳过备份")

            total = len(tasks)

            for idx, (start, end, name) in enumerate(tasks):
                self._update_status(f"正在拆分：{name}  ({idx+1}/{total})")
                self._update_progress((idx / total) * 100)

                # 创建新PDF
                safe_name = sanitize_filename(name)
                # 添加序号前缀确保排序正确
                out_filename = f"{idx+1:02d}_{safe_name}.pdf"
                out_path = os.path.join(output_dir, out_filename)

                new_doc = fitz.open()
                # fitz页码是0-based，用户输入是1-based
                new_doc.insert_pdf(self.pdf_doc, from_page=start - 1, to_page=end - 1)
                new_doc.save(out_path)
                new_doc.close()

                page_count = end - start + 1
                self._log(f"  已生成：{out_filename}  （第{start}-{end}页，共{page_count}页）")

            self._update_progress(100)
            self._update_status("拆分完成！")
            self._log(f"\n拆分完成！共生成 {total} 个文件。")
            self._log(f"输出位置：{output_dir}")

            self.root.after(0, self._on_complete)

        except Exception as e:
            self._log(f"\n拆分出错：{e}")
            self._update_status(f"出错：{e}")
            self.root.after(0, lambda: messagebox.showerror("拆分出错", f"拆分过程中出错：\n{e}"))

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.start_btn.configure(state=tk.NORMAL))

    def _on_complete(self):
        self.open_folder_btn.pack(side=tk.LEFT, padx=(10, 0))
        messagebox.showinfo("完成", f"拆分完成！\n\n共生成文件保存在：\n{self.output_dir}")

    # ================================================================
    #  辅助方法
    # ================================================================

    def _log(self, msg):
        def _do():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + '\n')
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        self.root.after(0, _do)

    def _update_status(self, text):
        self.root.after(0, lambda: self.status_label.configure(text=text))

    def _update_progress(self, value):
        self.root.after(0, lambda: self.progress.configure(value=value))

    def _open_output_folder(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            if sys.platform == 'win32':
                os.startfile(self.output_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', self.output_dir])
            else:
                subprocess.Popen(['xdg-open', self.output_dir])


# ================================================================
#  程序入口
# ================================================================

def main():
    if not fitz:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "缺少依赖",
            "运行此程序需要 PyMuPDF 库。\n\n"
            "请打开命令提示符（CMD），输入以下命令安装：\n\n"
            "    pip install PyMuPDF\n\n"
            "安装完成后重新运行本程序。"
        )
        sys.exit(1)

    if MODERN_UI:
        root = ttk.Window(themename="cosmo")
    else:
        root = tk.Tk()

    app = PDFSplitterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
