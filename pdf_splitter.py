#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF拆分工具 v1.0
支持按章节（书签）拆分和自定义页码范围拆分
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import shutil
import threading
import re
import subprocess

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def sanitize_filename(name):
    """移除文件名中的非法字符"""
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.strip(' .')
    return name if name else '未命名'


class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 拆分工具 v1.0")
        self.root.minsize(800, 700)

        # ---------- 状态变量 ----------
        self.pdf_path = None
        self.pdf_doc = None
        self.total_pages = 0
        self.chapters = []
        self.split_mode = tk.StringVar(value="chapter")
        self.chapter_vars = []          # [(BooleanVar, chapter_dict), ...]
        self.custom_rows = []           # [(frame, start_entry, end_entry, name_entry), ...]
        self.level_var = tk.StringVar(value="1")
        self.output_dir = None
        self.is_running = False

        # ---------- 样式 ----------
        style = ttk.Style()
        default_font = ('Microsoft YaHei UI', 10)
        style.configure('.', font=default_font)
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 18, 'bold'))
        style.configure('StepTitle.TLabel', font=('Microsoft YaHei UI', 11, 'bold'))
        style.configure('Big.TButton', font=('Microsoft YaHei UI', 11), padding=6)
        style.configure('Success.TLabel', foreground='green', font=('Microsoft YaHei UI', 10, 'bold'))
        style.configure('Info.TLabel', foreground='#555555', font=('Microsoft YaHei UI', 9))

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
        self.main_frame = ttk.Frame(canvas, padding=20)

        self.main_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.main_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮支持
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # 标题
        ttk.Label(self.main_frame, text="PDF 拆分工具", style='Title.TLabel').pack(pady=(0, 5))
        ttk.Label(self.main_frame, text="轻松将一本PDF拆分为多个小文件，原文件不会被修改",
                  style='Info.TLabel').pack(pady=(0, 15))

        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_log()

    # ---------- 第一步：选择文件 ----------

    def _build_step1(self):
        frame = ttk.LabelFrame(self.main_frame, text="  第一步：选择要拆分的PDF文件  ", padding=12)
        frame.pack(fill=tk.X, pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill=tk.X)

        ttk.Button(row, text="选择PDF文件 ...", command=self._select_file,
                   style='Big.TButton').pack(side=tk.LEFT)

        self.file_info_label = ttk.Label(row, text="  尚未选择文件", foreground='gray')
        self.file_info_label.pack(side=tk.LEFT, padx=(15, 0))

    # ---------- 第二步：选择拆分方式 ----------

    def _build_step2(self):
        frame = ttk.LabelFrame(self.main_frame, text="  第二步：选择拆分方式  ", padding=12)
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

        # 内容面板容器
        self.panel_container = ttk.Frame(frame)
        self.panel_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        # 章节面板
        self.chapter_panel = ttk.Frame(self.panel_container)
        self._build_chapter_panel()

        # 自定义面板
        self.custom_panel = ttk.Frame(self.panel_container)
        self._build_custom_panel()

        # 默认显示章节面板
        self.chapter_panel.pack(fill=tk.BOTH, expand=True)

    def _build_chapter_panel(self):
        # 层级选择
        top_row = ttk.Frame(self.chapter_panel)
        top_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(top_row, text="拆分层级：").pack(side=tk.LEFT)
        level_combo = ttk.Combobox(top_row, textvariable=self.level_var, width=25, state='readonly',
                                   values=["1", "2", "3"])
        level_combo.pack(side=tk.LEFT, padx=(4, 0))
        level_combo.bind('<<ComboboxSelected>>', lambda e: self._refresh_chapters())

        ttk.Label(top_row, text="（1=按章拆分，2=按节拆分，3=更细）",
                  style='Info.TLabel').pack(side=tk.LEFT, padx=(8, 0))

        # 按钮行
        btn_row = ttk.Frame(self.chapter_panel)
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_row, text="全选", command=lambda: self._toggle_all(True)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="取消全选", command=lambda: self._toggle_all(False)).pack(side=tk.LEFT)

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
        self.chapter_hint = ttk.Label(self.chapter_panel, text="请先选择一个PDF文件", foreground='gray')
        self.chapter_hint.pack(pady=8)

    def _build_custom_panel(self):
        hint = ttk.Label(self.custom_panel,
                         text="在下方添加要拆分的页码范围，\"名称\"可不填（将自动编号）。\n"
                              "页码从1开始，例如：第1页到第10页，就填  起始页=1  结束页=10",
                         style='Info.TLabel', justify=tk.LEFT)
        hint.pack(anchor=tk.W, pady=(0, 8))

        # 表头
        header = ttk.Frame(self.custom_panel)
        header.pack(fill=tk.X)
        ttk.Label(header, text="起始页", width=10, anchor=tk.CENTER,
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(header, text="结束页", width=10, anchor=tk.CENTER,
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(header, text="名称（可不填）", width=25, anchor=tk.CENTER,
                  font=('Microsoft YaHei UI', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 4))

        # 行容器
        self.custom_rows_frame = ttk.Frame(self.custom_panel)
        self.custom_rows_frame.pack(fill=tk.X, pady=(4, 0))

        # 添加按钮
        ttk.Button(self.custom_panel, text="+ 添加一行", command=self._add_range_row).pack(anchor=tk.W, pady=(6, 0))

        # 默认添加一行
        self._add_range_row()

    # ---------- 第三步：开始拆分 ----------

    def _build_step3(self):
        frame = ttk.LabelFrame(self.main_frame, text="  第三步：开始拆分  ", padding=12)
        frame.pack(fill=tk.X, pady=(0, 10))

        # 输出路径
        out_row = ttk.Frame(frame)
        out_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(out_row, text="输出位置：").pack(side=tk.LEFT)
        self.output_label = ttk.Label(out_row, text="（与源文件同一目录，自动创建文件夹）",
                                      foreground='gray')
        self.output_label.pack(side=tk.LEFT, padx=(4, 0))

        # 操作按钮
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)

        self.start_btn = ttk.Button(btn_row, text="开始拆分", command=self._start_split,
                                    style='Big.TButton')
        self.start_btn.pack(side=tk.LEFT)

        self.open_folder_btn = ttk.Button(btn_row, text="打开输出文件夹", command=self._open_output_folder,
                                          style='Big.TButton')
        self.open_folder_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.open_folder_btn.pack_forget()  # 初始隐藏

        # 进度条
        self.progress = ttk.Progressbar(frame, mode='determinate', length=400)
        self.progress.pack(fill=tk.X, pady=(10, 4))

        self.status_label = ttk.Label(frame, text="就绪", foreground='#555555')
        self.status_label.pack(anchor=tk.W)

    # ---------- 日志区域 ----------

    def _build_log(self):
        frame = ttk.LabelFrame(self.main_frame, text="  运行日志  ", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.log_text = tk.Text(frame, height=8, wrap=tk.WORD, font=('Consolas', 9),
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
                foreground='black'
            )

            self.output_dir = os.path.join(
                os.path.dirname(self.pdf_path),
                sanitize_filename(os.path.splitext(fname)[0]) + "_拆分结果"
            )
            self.output_label.configure(text=self.output_dir, foreground='black')

            self._refresh_chapters()
            self._log(f"文件加载成功，共 {self.total_pages} 页")

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
            self.chapter_hint.configure(text="此PDF没有书签/目录信息，请使用\"自定义页码范围拆分\"")
            return

        self.chapter_hint.configure(text=f"共检测到 {len(self.chapters)} 个章节，勾选要拆分的章节：")

        for i, ch in enumerate(self.chapters):
            var = tk.BooleanVar(value=True)
            row = ttk.Frame(self.chapter_inner)
            row.pack(fill=tk.X, padx=4, pady=1)

            cb = ttk.Checkbutton(row, variable=var)
            cb.pack(side=tk.LEFT)

            text = f"{ch['title']}    （第 {ch['start']} - {ch['end']} 页）"
            ttk.Label(row, text=text).pack(side=tk.LEFT, padx=(4, 0))

            self.chapter_vars.append((var, ch))

    def _toggle_all(self, state):
        for var, _ in self.chapter_vars:
            var.set(state)

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
                                command=lambda: self._remove_range_row(row_frame))
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

        if self.split_mode.get() == "chapter":
            self.chapter_panel.pack(fill=tk.BOTH, expand=True)
        else:
            self.custom_panel.pack(fill=tk.BOTH, expand=True)

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

    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
