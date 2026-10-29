# -*- coding: utf-8 -*-
"""
Word 操作题自动批改 — 图形界面
作者：Fantasy

视觉：中国试卷「装订线」构图——左侧朱红装订边 + 卷面纸白，
禁止蓝紫渐变；不用 SaaS 卡片堆叠。
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH,
    DISABLED,
    END,
    LEFT,
    NORMAL,
    RIGHT,
    WORD,
    BooleanVar,
    Canvas,
    Frame,
    Label,
    PhotoImage,
    StringVar,
    Tk,
    filedialog,
    messagebox,
    ttk,
)
from tkinter.scrolledtext import ScrolledText

# 保证同目录可导入
if getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(sys.executable).resolve().parent))
else:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from grade_exam import (  # noqa: E402
    AUTHOR,
    APP_TITLE,
    TEACHER_CLASSES,
    GradeRow,
    app_dir,
    deduction_stats,
    grade_all,
    load_typing_answer_text,
    precheck_submissions,
    regrade_one_in_rows,
    resource_dir,
    save_rows,
    self_check,
    suggest_output_path,
    unpack_submission_zip,
    unique_output_path,
)


# —— 设计令牌：教务室试卷，非通用仪表盘 ——
# paper 略偏冷灰绿，避开奶油 #F4F1EA；朱砂红作装订线，无蓝紫
C = {
    "desk": "#C8CFC2",  # 课桌贴面
    "paper": "#F1F0EA",  # 卷面
    "rule": "#D8D6CE",  # 细线
    "margin": "#B42318",  # 装订线朱红
    "margin_dark": "#8F1A12",
    "ink": "#1A1916",
    "muted": "#6E7168",
    "field": "#FFFFFF",
    "ok": "#2A6B4E",
    "warn": "#9A5B12",
    "err": "#B42318",
    "log_bg": "#EAE8E0",
    "log_fg": "#2A2E28",
}


def _pick_font(candidates: list[tuple], size: int, weight: str = "normal"):
    """按系统可用字体择优。"""
    # tk 字体名在运行时再绑定；此处仅返回元组，由调用方使用
    name = candidates[0]
    if weight == "bold":
        return (name, size, "bold")
    return (name, size)


class GradeApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(f"{APP_TITLE}  ·  {AUTHOR}")
        # 默认高度需盖住表单 +「批改记录」可见区域（高 DPI 下内容更高）
        self.root.geometry("920x860")
        self.root.minsize(800, 760)
        self.root.configure(bg=C["desk"])

        self.submissions = StringVar(value="")
        self.output = StringVar(value=str(suggest_output_path(app_dir())))
        self.class_mode = StringVar(value="auto")
        self.class_custom = StringVar(value="")
        self.skip_self_check = BooleanVar(value=False)
        self.busy = False
        self.last_output: Path | None = None
        self.last_rows: list[GradeRow] = []
        self.last_submissions: Path | None = None
        self.last_class_name: str | None = None
        self._answer_text: str | None = None
        self._review_win = None

        # 字体：标题用宋体（试卷抬头感），正文雅黑，日志等宽
        self.font_brand = _pick_font(["华文中宋", "STZhongsong", "SimSun", "宋体"], 22, "bold")
        self.font_brand_fallback = ("SimSun", 20, "bold")
        self.font_ui = ("Microsoft YaHei UI", 10)
        self.font_small = ("Microsoft YaHei UI", 9)
        self.font_btn = ("Microsoft YaHei UI", 11, "bold")
        self.font_mono = ("Consolas", 9)

        self._setup_style()
        self._apply_app_icon()
        self._build()

    def _asset_paths(self, *relative: str) -> list[Path]:
        rel = Path(*relative)
        bases = [resource_dir(), app_dir(), Path(__file__).resolve().parent]
        return [base / rel for base in bases]

    def _first_existing(self, paths: list[Path]) -> Path | None:
        for p in paths:
            if p.exists():
                return p
        return None

    def _apply_app_icon(self):
        """窗口标题栏 / 任务栏图标。"""
        self._icon_photos: list[PhotoImage] = []
        ico = self._first_existing(self._asset_paths("assets", "app.ico"))
        if ico is not None:
            try:
                self.root.iconbitmap(default=str(ico))
            except Exception:
                try:
                    self.root.iconbitmap(str(ico))
                except Exception:
                    pass
        # tk PhotoImage 对大图不稳定，优先 64px
        png = self._first_existing(
            self._asset_paths("assets", "app_icon_64.png")
            + self._asset_paths("assets", "app_icon.png")
        )
        if png is not None:
            try:
                img = PhotoImage(file=str(png))
                self.root.iconphoto(True, img)
                self._icon_photos.append(img)
            except Exception:
                pass

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Paper.TFrame", background=C["paper"])
        style.configure("Desk.TFrame", background=C["desk"])
        style.configure(
            "Ink.TLabel",
            background=C["paper"],
            foreground=C["ink"],
            font=self.font_ui,
        )
        style.configure(
            "Muted.TLabel",
            background=C["paper"],
            foreground=C["muted"],
            font=self.font_small,
        )
        style.configure(
            "Field.TEntry",
            fieldbackground=C["field"],
            foreground=C["ink"],
            insertcolor=C["ink"],
            bordercolor=C["rule"],
            lightcolor=C["rule"],
            darkcolor=C["rule"],
            padding=6,
        )
        style.configure(
            "Grade.TRadiobutton",
            background=C["paper"],
            foreground=C["ink"],
            font=self.font_ui,
            focuscolor=C["paper"],
        )
        style.configure(
            "Grade.TCheckbutton",
            background=C["paper"],
            foreground=C["ink"],
            font=self.font_ui,
            focuscolor=C["paper"],
        )
        style.map(
            "Grade.TRadiobutton",
            background=[("active", C["paper"])],
            foreground=[("active", C["ink"])],
        )
        style.map(
            "Grade.TCheckbutton",
            background=[("active", C["paper"])],
            foreground=[("active", C["ink"])],
        )
        style.configure(
            "Ghost.TButton",
            font=self.font_ui,
            padding=(12, 6),
            background=C["paper"],
            foreground=C["ink"],
        )
        style.configure(
            "Run.TButton",
            font=self.font_btn,
            padding=(22, 10),
            background=C["margin"],
            foreground="#FFFFFF",
            borderwidth=0,
        )
        style.map(
            "Run.TButton",
            background=[
                ("active", C["margin_dark"]),
                ("disabled", "#C9A8A4"),
            ],
            foreground=[("disabled", "#F5EDEC")],
        )
        style.map(
            "Ghost.TButton",
            background=[("active", C["rule"]), ("disabled", C["paper"])],
        )

    def _build(self):
        outer = Frame(self.root, bg=C["desk"])
        outer.pack(fill=BOTH, expand=True, padx=18, pady=18)

        # 整张「试卷」：左装订线 + 右卷面（单一构图）
        sheet = Frame(outer, bg=C["paper"], highlightthickness=0)
        sheet.pack(fill=BOTH, expand=True)

        # 左侧装订线（唯一强视觉锚点）
        bind = Canvas(sheet, width=36, bg=C["margin"], highlightthickness=0, bd=0)
        bind.pack(side=LEFT, fill="y")
        bind.bind("<Configure>", lambda e: self._draw_binding(bind, e.height))

        body = Frame(sheet, bg=C["paper"])
        body.pack(side=LEFT, fill=BOTH, expand=True)
        # 用 grid：日志行独占剩余高度，避免默认窗口把记录区挤成一条线
        body.columnconfigure(0, weight=1)
        body.rowconfigure(5, weight=1, minsize=180)

        # —— 卷首：品牌即主角 ——
        head = Frame(body, bg=C["paper"])
        head.grid(row=0, column=0, sticky="ew", padx=28, pady=(16, 4))

        brand = Label(
            head,
            text="OfficeGrade",
            bg=C["paper"],
            fg=C["ink"],
            font=self.font_brand,
            anchor="w",
        )
        brand.pack(anchor="w")
        try:
            brand.configure(font=self.font_brand)
        except Exception:
            brand.configure(font=self.font_brand_fallback)

        Label(
            head,
            text="选作业总文件夹（或 ZIP）→ 可先预检 → 一键批改 → 复核改分。成绩含预检/学情/教务简表。",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_ui,
            anchor="w",
            justify=LEFT,
            wraplength=720,
        ).pack(anchor="w", pady=(4, 0))

        Label(
            head,
            text=f"作者 {AUTHOR}　　支持 .docx / .doc",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_small,
            anchor="w",
        ).pack(anchor="w", pady=(2, 0))

        Canvas(body, height=1, bg=C["rule"], highlightthickness=0, bd=0).grid(
            row=1, column=0, sticky="ew", padx=28, pady=(8, 6)
        )

        form = Frame(body, bg=C["paper"])
        form.grid(row=2, column=0, sticky="ew", padx=28)

        self._field_block(
            form,
            "学生作业",
            "点「浏览」可选文件夹，或 ZIP 压缩包（自动解压）",
            self.submissions,
            self._browse_submissions,
        )
        self._field_block(
            form,
            "成绩表保存为",
            None,
            self.output,
            self._browse_output,
        )

        Label(
            form, text="成绩单「班级」列", bg=C["paper"], fg=C["ink"], font=self.font_ui
        ).pack(anchor="w", pady=(6, 2))
        opts = Frame(form, bg=C["paper"])
        opts.pack(fill="x")
        ttk.Radiobutton(
            opts,
            text="自动（文件夹名里出现最多的班级）",
            variable=self.class_mode,
            value="auto",
            command=self._toggle_class,
            style="Grade.TRadiobutton",
        ).pack(anchor="w")
        ttk.Radiobutton(
            opts,
            text="留空",
            variable=self.class_mode,
            value="empty",
            command=self._toggle_class,
            style="Grade.TRadiobutton",
        ).pack(anchor="w")

        custom_row = Frame(form, bg=C["paper"])
        custom_row.pack(fill="x", pady=(2, 0))
        ttk.Radiobutton(
            custom_row,
            text="指定",
            variable=self.class_mode,
            value="custom",
            command=self._toggle_class,
            style="Grade.TRadiobutton",
        ).pack(side=LEFT)
        self.class_entry = ttk.Entry(
            custom_row,
            textvariable=self.class_custom,
            font=self.font_ui,
            style="Field.TEntry",
            width=34,
        )
        self.class_entry.pack(side=LEFT, padx=(8, 0))
        # 点输入框即切到「指定」并获焦（禁用态点不进，故未选中时用只读可点）
        self.class_entry.bind("<Button-1>", self._activate_custom_class)
        self.class_entry.bind("<FocusIn>", self._activate_custom_class)

        Label(
            form,
            text="任课班参考：" + "、".join(TEACHER_CLASSES[:4]) + "…",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_small,
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        ttk.Checkbutton(
            form,
            text="跳过答案自检（加快速度）",
            variable=self.skip_self_check,
            style="Grade.TCheckbutton",
        ).pack(anchor="w", pady=(4, 0))

        actions = Frame(body, bg=C["paper"])
        actions.grid(row=3, column=0, sticky="ew", padx=28, pady=(12, 6))

        self.btn_run = ttk.Button(
            actions, text="开始批改", style="Run.TButton", command=self._start
        )
        self.btn_run.pack(side=LEFT)

        self.btn_precheck = ttk.Button(
            actions,
            text="预检",
            style="Ghost.TButton",
            command=self._run_precheck,
        )
        self.btn_precheck.pack(side=LEFT, padx=(10, 0))

        self.btn_open = ttk.Button(
            actions,
            text="打开成绩表",
            style="Ghost.TButton",
            command=self._open_excel,
            state=DISABLED,
        )
        self.btn_open.pack(side=LEFT, padx=(10, 0))
        self.btn_folder = ttk.Button(
            actions,
            text="打开输出目录",
            style="Ghost.TButton",
            command=self._open_outdir,
            state=DISABLED,
        )
        self.btn_folder.pack(side=LEFT, padx=(8, 0))
        self.btn_review = ttk.Button(
            actions,
            text="复核改分",
            style="Ghost.TButton",
            command=self._open_review,
            state=DISABLED,
        )
        self.btn_review.pack(side=LEFT, padx=(8, 0))

        self.status = Label(
            actions,
            text="就绪",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_ui,
        )
        self.status.pack(side=RIGHT)

        Canvas(body, height=1, bg=C["rule"], highlightthickness=0, bd=0).grid(
            row=4, column=0, sticky="ew", padx=28, pady=(2, 4)
        )

        log_wrap = Frame(body, bg=C["paper"])
        log_wrap.grid(row=5, column=0, sticky="nsew", padx=28, pady=(0, 14))
        log_wrap.rowconfigure(1, weight=1)
        log_wrap.columnconfigure(0, weight=1)
        Label(
            log_wrap,
            text="批改记录",
            bg=C["paper"],
            fg=C["ink"],
            font=self.font_ui,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.log = ScrolledText(
            log_wrap,
            height=10,
            wrap=WORD,
            font=self.font_mono,
            bg=C["log_bg"],
            fg=C["log_fg"],
            insertbackground=C["ink"],
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=C["rule"],
            highlightcolor=C["margin"],
            padx=10,
            pady=8,
        )
        self.log.grid(row=1, column=0, sticky="nsew")
        self._log(f"{APP_TITLE}\n作者：{AUTHOR}\n程序目录：{app_dir()}\n")
        self._toggle_class()
        self.root.after_idle(self._place_on_screen)

    def _work_area(self) -> tuple[int, int, int, int]:
        """返回主显示器工作区 (left, top, width, height)，已扣除任务栏。"""
        try:
            from ctypes import byref, windll, wintypes

            rect = wintypes.RECT()
            # SPI_GETWORKAREA = 0x0030；必须 byref，否则可能得到空矩形
            ok = windll.user32.SystemParametersInfoW(0x0030, 0, byref(rect), 0)
            if ok:
                w = int(rect.right - rect.left)
                h = int(rect.bottom - rect.top)
                if w >= 400 and h >= 300:
                    return int(rect.left), int(rect.top), w, h
        except Exception:
            pass
        sw = max(self.root.winfo_screenwidth(), 800)
        sh = max(self.root.winfo_screenheight(), 600)
        return 0, 0, sw, sh

    def _place_on_screen(self):
        """按内容抬高窗口，并居中到工作区内，避免启动时贴边/被任务栏遮挡。"""
        self.root.update_idletasks()
        left, top, work_w, work_h = self._work_area()
        if work_w < 400 or work_h < 300:
            left, top, work_w, work_h = 0, 0, 1280, 800

        margin = 24
        max_w = max(640, work_w - margin * 2)
        max_h = max(560, work_h - margin * 2)

        need_w = min(max(self.root.winfo_reqwidth(), 920), max_w)
        need_h = min(max(860, self.root.winfo_reqheight() + 24), max_h)
        need_h = max(min(need_h, max_h), min(760, max_h))
        need_w = max(min(need_w, max_w), min(800, max_w))

        x = left + max(0, (work_w - need_w) // 2)
        y = top + max(0, (work_h - need_h) // 2)
        # 保证标题栏落在工作区内
        max_x = left + max(0, work_w - need_w)
        max_y = top + max(0, work_h - need_h)
        x = min(max(x, left), max_x)
        y = min(max(y, top), max_y)
        self.root.geometry(f"{need_w}x{need_h}+{x}+{y}")
        self.root.minsize(min(800, max_w), min(760, max_h))
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def _draw_binding(self, canvas: Canvas, height: int):
        canvas.delete("all")
        w = int(canvas["width"])
        canvas.create_rectangle(0, 0, w, height, fill=C["margin"], outline="")
        # 装订孔：卷面记忆点
        y = 28
        while y < height - 20:
            canvas.create_oval(
                11, y, 25, y + 14, fill=C["margin_dark"], outline=C["margin_dark"]
            )
            canvas.create_oval(13, y + 2, 23, y + 12, fill=C["paper"], outline="")
            y += 48
        # 竖排小字
        canvas.create_text(
            w // 2,
            min(height - 40, height // 2 + 120),
            text="装\n订\n线",
            fill="#F7D4CF",
            font=("Microsoft YaHei UI", 8),
            justify="center",
        )

    def _field_block(self, parent, title, hint, var, browse_cmd):
        Label(parent, text=title, bg=C["paper"], fg=C["ink"], font=self.font_ui).pack(
            anchor="w", pady=(4, 1)
        )
        if hint:
            Label(
                parent, text=hint, bg=C["paper"], fg=C["muted"], font=self.font_small
            ).pack(anchor="w")
        row = Frame(parent, bg=C["paper"])
        row.pack(fill="x", pady=(2, 0))
        entry = ttk.Entry(row, textvariable=var, font=self.font_ui, style="Field.TEntry")
        entry.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(row, text="浏览", style="Ghost.TButton", command=browse_cmd).pack(
            side=RIGHT, padx=(8, 0)
        )

    def _toggle_class(self):
        # 未选「指定」时用 readonly：仍可点击获焦，但不能误改；选中后可编辑
        if self.class_mode.get() == "custom":
            self.class_entry.configure(state=NORMAL)
        else:
            self.class_entry.configure(state="readonly")

    def _activate_custom_class(self, _event=None):
        if self.class_mode.get() != "custom":
            self.class_mode.set("custom")
            self._toggle_class()
        # 推迟获焦，避免与 readonly→normal 切换抢事件
        self.root.after_idle(lambda: self.class_entry.focus_set())

    def _browse_submissions(self):
        """浏览：可选文件夹或 ZIP（ZIP 会自动解压并填入解压目录）。"""
        choice = self._ask_folder_or_zip()
        if choice is None:
            return
        initial = str(app_dir())
        if choice == "folder":
            path = filedialog.askdirectory(
                title="选择学生作业总文件夹（一人一个子文件夹）",
                initialdir=initial,
            )
            if not path:
                return
            self._apply_submissions_path(Path(path), quiet=True)
            return

        zpath = filedialog.askopenfilename(
            title="选择学生作业 ZIP 压缩包",
            filetypes=[("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
            initialdir=initial,
        )
        if not zpath:
            return
        self._apply_submissions_path(Path(zpath), quiet=False)

    def _ask_folder_or_zip(self) -> str | None:
        """返回 'folder' / 'zip' / None(取消)。"""
        from tkinter import Toplevel

        top = Toplevel(self.root)
        top.title("选择作业来源")
        top.configure(bg=C["paper"])
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()
        result: dict[str, str | None] = {"v": None}

        Label(
            top,
            text="学生作业是文件夹，还是 ZIP 压缩包？",
            bg=C["paper"],
            fg=C["ink"],
            font=self.font_ui,
            anchor="w",
        ).pack(fill="x", padx=20, pady=(18, 8))
        Label(
            top,
            text="选 ZIP 后会自动解压到同级「原名_解压」目录。",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_small,
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 12))

        bar = Frame(top, bg=C["paper"])
        bar.pack(fill="x", padx=20, pady=(0, 18))
        for col in range(3):
            bar.columnconfigure(col, weight=1, uniform="choice")

        def pick(v: str):
            result["v"] = v
            top.destroy()

        # 三钮等宽等高：统一 padding；主选项仅用朱红底区分
        style = ttk.Style()
        style.configure(
            "Choice.TButton",
            font=self.font_ui,
            padding=(12, 8),
            background=C["margin"],
            foreground="#FFFFFF",
            borderwidth=0,
        )
        style.map(
            "Choice.TButton",
            background=[("active", C["margin_dark"]), ("disabled", "#C9A8A4")],
            foreground=[("disabled", "#F5EDEC")],
        )
        style.configure(
            "ChoiceGhost.TButton",
            font=self.font_ui,
            padding=(12, 8),
            background=C["paper"],
            foreground=C["ink"],
        )
        style.map(
            "ChoiceGhost.TButton",
            background=[("active", C["rule"]), ("disabled", C["paper"])],
        )

        btn_w = 12
        ttk.Button(
            bar,
            text="选文件夹",
            style="Choice.TButton",
            width=btn_w,
            command=lambda: pick("folder"),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            bar,
            text="选 ZIP 包",
            style="ChoiceGhost.TButton",
            width=btn_w,
            command=lambda: pick("zip"),
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(
            bar,
            text="取消",
            style="ChoiceGhost.TButton",
            width=btn_w,
            command=lambda: pick(""),
        ).grid(row=0, column=2, sticky="ew")

        top.update_idletasks()
        # 固定弹窗宽度，三列均分后视觉更齐
        min_w = max(top.winfo_reqwidth(), 420)
        top.minsize(min_w, top.winfo_reqheight())
        tw, th = min_w, top.winfo_reqheight()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        x = rx + max(0, (rw - tw) // 2)
        y = ry + max(0, (rh - th) // 2)
        top.geometry(f"{tw}x{th}+{x}+{y}")
        top.wait_window()
        v = result["v"]
        if not v:
            return None
        return v

    def _apply_submissions_path(self, path: Path, quiet: bool = True) -> Path | None:
        """把路径写入作业栏；若是 ZIP 则先解压。成功返回最终目录。"""
        path = Path(path).expanduser()
        try:
            path = path.resolve(strict=False)
        except Exception:
            pass
        if not path.exists():
            messagebox.showwarning("提示", f"路径不存在：\n{path}")
            return None

        dest = path
        if path.is_file():
            if path.suffix.lower() != ".zip":
                messagebox.showwarning("提示", "请选择文件夹，或 .zip 压缩包。")
                return None
            try:
                self._set_status("正在解压 ZIP…", "warn")
                self.root.update_idletasks()
                dest = unpack_submission_zip(path)
            except Exception as e:
                self._set_status("解压失败", "err")
                messagebox.showerror("解压失败", str(e))
                return None
            self._log(f"已解压 ZIP「{path.name}」→ {dest}")
            if not quiet:
                messagebox.showinfo(
                    "解压完成",
                    f"已解压到：\n{dest}\n\n已填入作业路径，可先「预检」或直接「开始批改」。",
                )
            self._set_status("就绪", "muted")
        elif not path.is_dir():
            messagebox.showwarning("提示", "请选择文件夹或 ZIP 压缩包。")
            return None

        self.submissions.set(str(dest))
        hint = self.class_custom.get().strip() if self.class_mode.get() == "custom" else ""
        self.output.set(str(suggest_output_path(app_dir(), hint)))
        return dest

    def _resolve_submissions_dir(self) -> Path | None:
        """读取界面路径；若仍是 ZIP 则解压。无效则提示并返回 None。"""
        raw_sub = self.submissions.get().strip().strip('"')
        if not raw_sub or raw_sub in {".", "./", ".\\"}:
            messagebox.showwarning(
                "提示",
                "请先选择学生作业（点「浏览」可选文件夹或 ZIP 包）。",
            )
            return None
        sub = Path(raw_sub).expanduser()
        try:
            sub = sub.resolve(strict=False)
        except Exception:
            pass
        if not sub.exists():
            messagebox.showwarning("提示", f"路径不存在：\n{sub}")
            return None
        if sub.is_file() and sub.suffix.lower() == ".zip":
            return self._apply_submissions_path(sub, quiet=False)
        if sub.is_dir():
            return sub
        messagebox.showwarning("提示", "请选择有效的作业文件夹，或 .zip 压缩包。")
        return None

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="成绩表保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialfile=Path(self.output.get() or "成绩表.xlsx").name,
            initialdir=str(app_dir()),
        )
        if path:
            self.output.set(path)

    def _run_precheck(self):
        if self.busy:
            return
        sub = self._resolve_submissions_dir()
        if sub is None:
            return

        self.busy = True
        self.btn_run.configure(state=DISABLED)
        self.btn_precheck.configure(state=DISABLED)
        self._set_status("预检中…", "warn")

        def work():
            try:
                _students, pre = precheck_submissions(sub)
                ok = sum(1 for p in pre if p.status == "可批改")
                warn = [p for p in pre if p.can_grade and p.status != "可批改"]
                bad = [p for p in pre if not p.can_grade]
                lines = [
                    f"预检完成：共 {len(pre)} 人",
                    f"可批改 {ok}　部分问题 {len(warn)}　无法批改 {len(bad)}",
                ]
                for p in warn + bad:
                    lines.append(f"· {p.student_name or p.folder}：{p.status}｜{p.issues}")
                text = "\n".join(lines)
                self.root.after(0, lambda: self._precheck_done(text, len(bad)))
            except Exception as e:
                detail = traceback.format_exc()
                self.root.after(
                    0,
                    lambda: self._done_err(self._friendly_error(e), detail),
                )

        threading.Thread(target=work, daemon=True).start()

    def _precheck_done(self, text: str, bad_n: int):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self.btn_precheck.configure(state=NORMAL)
        self._log("\n" + "—" * 36)
        self._log(text)
        self._set_status("预检完成", "warn" if bad_n else "ok")
        messagebox.showinfo("预检结果", text if len(text) < 1200 else text[:1200] + "\n…")

    def _log(self, msg: str):
        self.log.configure(state=NORMAL)
        self.log.insert(END, msg + ("" if msg.endswith("\n") else "\n"))
        self.log.see(END)
        self.log.configure(state=DISABLED)

    def _set_status(self, text: str, tone: str = "muted"):
        color = {"ok": C["ok"], "err": C["err"], "warn": C["warn"]}.get(tone, C["muted"])
        self.status.configure(text=text, fg=color)

    def _start(self):
        if self.busy:
            return
        sub = self._resolve_submissions_dir()
        if sub is None:
            return
        raw_out = self.output.get().strip().strip('"')
        if not raw_out or raw_out in {".", "./", ".\\"}:
            messagebox.showwarning(
                "提示", "请填写成绩表保存路径，或点右侧「浏览」选择保存位置。"
            )
            return
        out = Path(raw_out).expanduser()
        try:
            out = out.resolve(strict=False)
        except Exception:
            pass
        if out.exists() and out.is_dir():
            out = suggest_output_path(out)
        if out.suffix.lower() not in {".xlsx", ".xlsm"}:
            messagebox.showwarning("提示", "成绩表请保存为 .xlsx 文件。")
            return
        if out.parent == out or not out.parent.exists():
            messagebox.showwarning("提示", "成绩表保存目录不存在，请重新选择保存位置。")
            return
        # 已存在则自动换名，避免覆盖上一份
        out = unique_output_path(out)
        self.output.set(str(out))

        mode = self.class_mode.get()
        if mode == "empty":
            class_name = ""
        elif mode == "custom":
            class_name = self.class_custom.get().strip()
            if not class_name:
                messagebox.showwarning("提示", "请填写指定的班级名称，或改选「自动/留空」。")
                return
        else:
            class_name = None

        self.busy = True
        self.btn_run.configure(state=DISABLED)
        self.btn_precheck.configure(state=DISABLED)
        self.btn_open.configure(state=DISABLED)
        self.btn_folder.configure(state=DISABLED)
        self.btn_review.configure(state=DISABLED)
        self._set_status("批改中…", "warn")
        self._log("\n" + "—" * 36)
        self._log(f"提交目录：{sub}")
        self._log(f"输出文件：{out}")
        if class_name is None:
            self._log("班级：自动识别")
        else:
            self._log(f"班级：{class_name if class_name else '（空）'}")

        skip = self.skip_self_check.get()

        def work():
            try:
                if not skip:
                    self.root.after(0, lambda: self._log("正在自检标准答案…"))
                    self_check()
                    self.root.after(0, lambda: self._log("自检通过。"))

                def on_progress(i, total, name):
                    self.root.after(
                        0,
                        lambda i=i, total=total, name=name: self._log(
                            f"[{i}/{total}] {name}"
                        ),
                    )
                    self.root.after(
                        0,
                        lambda i=i, total=total: self._set_status(
                            f"批改中 {i}/{total}", "warn"
                        ),
                    )

                rows = grade_all(
                    sub,
                    out,
                    class_name=class_name,
                    on_progress=on_progress,
                    avoid_overwrite=True,
                )
                real_out = getattr(grade_all, "last_output", out)
                totals = [r.effective_total() for r in rows]
                used = rows[0].class_name if rows else ""
                summary = (
                    f"完成：共 {len(rows)} 人。\n"
                    f"班级栏：{used if used else '（空）'}\n"
                    f"输出：{real_out}"
                )
                if totals:
                    summary += (
                        f"\n总分 平均 {sum(totals)/len(totals):.1f}，"
                        f"最高 {max(totals)}，最低 {min(totals)}"
                    )
                top = deduction_stats(rows)[:3]
                if top:
                    summary += "\n高频未满分：" + "；".join(
                        f"{a}({b}人)" for a, b, _ in top
                    )
                self.root.after(
                    0,
                    lambda: self._done_ok(real_out, summary, rows, sub, class_name),
                )
            except Exception as e:
                detail = traceback.format_exc()
                user_msg = self._friendly_error(e)
                self.root.after(
                    0, lambda user_msg=user_msg, detail=detail: self._done_err(user_msg, detail)
                )

        threading.Thread(target=work, daemon=True).start()

    def _friendly_error(self, e: BaseException) -> str:
        """给教师看的短说明；技术堆栈只写进批改记录。"""
        if isinstance(e, PermissionError):
            return (
                "无法写入成绩表。\n"
                "请确认保存路径有效，且该 Excel 没有正在被打开。"
            )
        if isinstance(e, FileNotFoundError):
            return str(e) or "找不到需要的文件或文件夹。"
        if isinstance(e, OSError):
            return f"保存或读取文件时出错：{e}"
        msg = str(e).strip()
        if msg:
            return f"批改未能完成：{msg}"
        return "批改未能完成，请查看下方「批改记录」。"

    def _done_ok(
        self,
        out: Path,
        summary: str,
        rows: list[GradeRow] | None = None,
        submissions: Path | None = None,
        class_name: str | None = None,
    ):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self.btn_precheck.configure(state=NORMAL)
        self.btn_open.configure(state=NORMAL)
        self.btn_folder.configure(state=NORMAL)
        self.btn_review.configure(state=NORMAL)
        self.last_output = out
        self.output.set(str(out))
        if rows is not None:
            self.last_rows = rows
        if submissions is not None:
            self.last_submissions = submissions
        self.last_class_name = class_name
        self._log(summary)
        self._set_status("完成", "ok")
        if messagebox.askyesno(
            "批改完成",
            summary + "\n\n是否打开「复核改分」窗口？（可打开学生文件、改分、单人重批）",
        ):
            self._open_review()

    def _done_err(self, user_msg: str, detail: str = ""):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self.btn_precheck.configure(state=NORMAL)
        if detail:
            self._log("出错了：\n" + detail)
        else:
            self._log("出错了：\n" + user_msg)
        self._set_status("出错", "err")
        messagebox.showerror("批改失败", user_msg)

    def _open_excel(self):
        path = self.last_output or Path(self.output.get())
        if path.exists():
            os.startfile(str(path))  # noqa: S606
        else:
            messagebox.showwarning("提示", "还没有生成成绩表。")

    def _open_outdir(self):
        path = self.last_output or Path(self.output.get())
        folder = path.parent if path.suffix else path
        if folder.exists():
            os.startfile(str(folder))  # noqa: S606

    def _selected_row(self, tree) -> GradeRow | None:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先在列表里点选一名学生。")
            return None
        idx = int(sel[0])
        if idx < 0 or idx >= len(self.last_rows):
            return None
        return self.last_rows[idx]

    def _open_path(self, path_str: str, missing_msg: str):
        if not path_str:
            messagebox.showwarning("提示", missing_msg)
            return
        p = Path(path_str)
        if not p.exists():
            messagebox.showwarning("提示", f"文件不存在：\n{p}")
            return
        os.startfile(str(p))  # noqa: S606

    def _persist_rows(self, note: str = ""):
        if not self.last_output:
            messagebox.showwarning("提示", "还没有成绩表路径。")
            return
        try:
            out = save_rows(
                self.last_rows,
                self.last_output,
                submissions=self.last_submissions,
                avoid_overwrite=False,
            )
            self.last_output = out
            self.output.set(str(out))
            if note:
                self._log(note)
            self._log(f"已写回：{out}")
        except Exception as e:
            messagebox.showerror("保存失败", self._friendly_error(e))
            self._log(traceback.format_exc())

    def _open_review(self):
        if not self.last_rows:
            messagebox.showwarning("提示", "请先完成一次批改，再打开复核。")
            return

        if self._review_win is not None:
            try:
                if self._review_win.winfo_exists():
                    self._review_win.lift()
                    self._review_win.focus_force()
                    return
            except Exception:
                pass

        from tkinter import Toplevel, simpledialog

        top = Toplevel(self.root)
        top.title("复核改分 · OfficeGrade")
        top.configure(bg=C["paper"])
        self._review_win = top
        top.protocol("WM_DELETE_WINDOW", lambda: self._close_review(top))
        # 相对主窗口居中
        win_w, win_h = 920, 560
        self.root.update_idletasks()
        rw = max(self.root.winfo_width(), 1)
        rh = max(self.root.winfo_height(), 1)
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        x = rx + max(0, (rw - win_w) // 2)
        y = ry + max(0, (rh - win_h) // 2)
        top.geometry(f"{win_w}x{win_h}+{x}+{y}")
        top.transient(self.root)
        top.lift()
        top.focus_force()

        Label(
            top,
            text="点选学生 → 打开打字/简历核对 → 必要时改最终分或单人重批。改完会写回当前成绩表。",
            bg=C["paper"],
            fg=C["muted"],
            font=self.font_small,
            anchor="w",
            justify=LEFT,
            wraplength=880,
        ).pack(fill="x", padx=12, pady=(10, 4))

        cols = ("姓名", "自动分", "最终分", "打字", "简历", "文件夹")
        tree = ttk.Treeview(top, columns=cols, show="headings", height=18)
        for c, w in zip(cols, (90, 80, 80, 60, 60, 420)):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor="center" if c != "文件夹" else "w")
        tree.pack(fill=BOTH, expand=True, padx=12, pady=6)

        def refill():
            tree.delete(*tree.get_children())
            for i, r in enumerate(self.last_rows):
                tree.insert(
                    "",
                    END,
                    iid=str(i),
                    values=(
                        r.student_name,
                        r.total,
                        r.effective_total(),
                        "有" if r.typing_file else "缺",
                        "有" if r.resume_file else "缺",
                        r.folder,
                    ),
                )

        refill()
        self._review_tree = tree
        self._review_refill = refill

        bar = Frame(top, bg=C["paper"])
        bar.pack(fill="x", padx=12, pady=(0, 12))

        def open_typing():
            r = self._selected_row(tree)
            if r:
                self._open_path(r.typing_fullpath, "该生没有打字文件路径。")

        def open_resume():
            r = self._selected_row(tree)
            if r:
                self._open_path(r.resume_fullpath, "该生没有简历文件路径。")

        def open_folder():
            r = self._selected_row(tree)
            if r:
                self._open_path(r.folder_path, "找不到该生文件夹。")

        def adjust_score():
            r = self._selected_row(tree)
            if not r:
                return
            cur = r.effective_total()
            raw = simpledialog.askstring(
                "改最终总分",
                f"{r.student_name or r.folder}\n自动分 {r.total}，当前最终分 {cur}\n请输入新的最终总分（0–100）：",
                parent=top,
                initialvalue=str(cur),
            )
            if raw is None:
                return
            try:
                val = float(raw.strip())
            except ValueError:
                messagebox.showwarning("提示", "请输入数字。", parent=top)
                return
            if val < 0 or val > 100:
                messagebox.showwarning("提示", "分数需在 0–100。", parent=top)
                return
            reason = simpledialog.askstring(
                "改分说明",
                "请写一句改分原因（会写入成绩表「改分说明」列）：",
                parent=top,
                initialvalue=r.adjust_note or "人工复核调整",
            )
            if reason is None:
                return
            r.final_total = round(val, 1)
            r.adjust_note = (reason or "").strip()
            refill()
            self._persist_rows(
                f"已改分：{r.student_name or r.folder} → 最终 {r.final_total}（{r.adjust_note}）"
            )
            messagebox.showinfo("已保存", "最终总分已写回成绩表。", parent=top)

        def regrade_one():
            r = self._selected_row(tree)
            if not r:
                return
            if not self.last_submissions:
                messagebox.showwarning("提示", "缺少提交目录，无法重批。", parent=top)
                return
            if not messagebox.askyesno(
                "单人重批",
                f"确认重新批改「{r.student_name or r.folder}」？\n会按当前文件夹里的文件重算自动分。",
                parent=top,
            ):
                return
            try:
                if self._answer_text is None:
                    self._answer_text = load_typing_answer_text()
                self.last_rows = regrade_one_in_rows(
                    self.last_rows,
                    r.folder,
                    self.last_submissions,
                    answer_text=self._answer_text,
                    class_name=r.class_name,
                )
                refill()
                self._persist_rows(f"已重批：{r.student_name or r.folder}")
                messagebox.showinfo("已重批", "该生已重批并写回成绩表。", parent=top)
            except Exception as e:
                messagebox.showerror("重批失败", self._friendly_error(e), parent=top)
                self._log(traceback.format_exc())

        ttk.Button(bar, text="打开打字", style="Ghost.TButton", command=open_typing).pack(
            side=LEFT
        )
        ttk.Button(bar, text="打开简历", style="Ghost.TButton", command=open_resume).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(bar, text="打开文件夹", style="Ghost.TButton", command=open_folder).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(bar, text="改最终分", style="Ghost.TButton", command=adjust_score).pack(
            side=LEFT, padx=(16, 0)
        )
        ttk.Button(bar, text="单人重批", style="Ghost.TButton", command=regrade_one).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(
            bar, text="打开成绩表", style="Ghost.TButton", command=self._open_excel
        ).pack(side=RIGHT)

    def _close_review(self, top):
        try:
            top.destroy()
        except Exception:
            pass
        self._review_win = None


def run_gui():
    root = Tk()
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    GradeApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()
