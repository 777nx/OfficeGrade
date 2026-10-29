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
    app_dir,
    grade_all,
    resource_dir,
    self_check,
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
        self.output = StringVar(value=str(app_dir() / "成绩表.xlsx"))
        self.class_mode = StringVar(value="auto")
        self.class_custom = StringVar(value="")
        self.skip_self_check = BooleanVar(value=False)
        self.busy = False
        self.last_output: Path | None = None

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
            text="把学生作业总文件夹选好，一键批改打字、简历表和命名，导出 Excel。",
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
            "学生作业文件夹",
            "选「一人一个子文件夹」的总目录",
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
        path = filedialog.askdirectory(
            title="选择学生作业总文件夹",
            initialdir=str(app_dir()),
        )
        if path:
            self.submissions.set(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="成绩表保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialfile="成绩表.xlsx",
            initialdir=str(app_dir()),
        )
        if path:
            self.output.set(path)

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
        raw_sub = self.submissions.get().strip().strip('"')
        # Path("") 在 Windows 上等于「当前目录」，会误批 release/答案
        if not raw_sub or raw_sub in {".", "./", ".\\"}:
            messagebox.showwarning("提示", "请先选择学生作业文件夹（浏览选「一人一个子文件夹」的总目录）。")
            return
        sub = Path(raw_sub).expanduser()
        try:
            sub = sub.resolve(strict=False)
        except Exception:
            pass
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
            out = out / "成绩表.xlsx"
        if out.suffix.lower() not in {".xlsx", ".xlsm"}:
            messagebox.showwarning("提示", "成绩表请保存为 .xlsx 文件。")
            return
        if not sub.exists() or not sub.is_dir():
            messagebox.showwarning("提示", "请先选择有效的学生作业文件夹。")
            return
        if out.parent == out or not out.parent.exists():
            messagebox.showwarning("提示", "成绩表保存目录不存在，请重新选择保存位置。")
            return
        # 同步回界面，避免空路径再次触发
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
        self.btn_open.configure(state=DISABLED)
        self.btn_folder.configure(state=DISABLED)
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
                    sub, out, class_name=class_name, on_progress=on_progress
                )
                totals = [r.total for r in rows]
                used = rows[0].class_name if rows else ""
                summary = (
                    f"完成：共 {len(rows)} 人。\n"
                    f"班级栏：{used if used else '（空）'}\n"
                    f"输出：{out}"
                )
                if totals:
                    summary += (
                        f"\n总分 平均 {sum(totals)/len(totals):.1f}，"
                        f"最高 {max(totals)}，最低 {min(totals)}"
                    )
                self.root.after(0, lambda: self._done_ok(out, summary))
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

    def _done_ok(self, out: Path, summary: str):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self.btn_open.configure(state=NORMAL)
        self.btn_folder.configure(state=NORMAL)
        self.last_output = out
        self._log(summary)
        self._set_status("完成", "ok")
        messagebox.showinfo("批改完成", summary)

    def _done_err(self, user_msg: str, detail: str = ""):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
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
