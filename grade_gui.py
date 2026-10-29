# -*- coding: utf-8 -*-
"""
Word 操作题自动批改 — 图形界面
作者：Fantasy
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
    self_check,
)


# 视觉：粉笔黑板感的教务工具，避开默认紫/奶油模板
COLORS = {
    "bg": "#1B2A2F",
    "panel": "#243B42",
    "card": "#2E4A52",
    "accent": "#E8A87C",
    "accent2": "#6BBFBF",
    "text": "#F4F1EC",
    "muted": "#A8B5B8",
    "ok": "#7DCEA0",
    "err": "#E08A8A",
    "input_bg": "#152226",
}


class GradeApp:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(f"{APP_TITLE}  ·  {AUTHOR}")
        self.root.geometry("820x640")
        self.root.minsize(720, 560)
        self.root.configure(bg=COLORS["bg"])

        self.submissions = StringVar(value="")
        self.output = StringVar(value=str(app_dir() / "成绩表.xlsx"))
        self.class_mode = StringVar(value="auto")  # auto / empty / custom
        self.class_custom = StringVar(value="")
        self.skip_self_check = BooleanVar(value=False)
        self.busy = False
        self.last_output: Path | None = None

        self._setup_style()
        self._build()

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        font_ui = ("Microsoft YaHei UI", 10)
        font_title = ("Microsoft YaHei UI", 18, "bold")
        self.font_ui = font_ui
        self.font_title = font_title
        self.font_mono = ("Consolas", 9)

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["panel"])
        style.configure(
            "Title.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=font_title,
        )
        style.configure(
            "Sub.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Card.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=font_ui,
        )
        style.configure(
            "Muted.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["muted"],
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Accent.TButton",
            font=("Microsoft YaHei UI", 11, "bold"),
            padding=(16, 8),
        )
        style.configure("TButton", font=font_ui, padding=(10, 6))
        style.configure(
            "TRadiobutton",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=font_ui,
        )
        style.configure(
            "TCheckbutton",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=font_ui,
        )
        style.map(
            "TRadiobutton",
            background=[("active", COLORS["panel"])],
            foreground=[("active", COLORS["text"])],
        )
        style.map(
            "TCheckbutton",
            background=[("active", COLORS["panel"])],
            foreground=[("active", COLORS["text"])],
        )

    def _build(self):
        pad = {"padx": 20, "pady": 8}

        head = ttk.Frame(self.root, style="TFrame")
        head.pack(fill="x", **pad)
        ttk.Label(head, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            head,
            text=f"作者 {AUTHOR}  ·  支持 .docx / .doc  ·  自动生成 Excel 成绩表",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=20, pady=6)

        # 提交目录
        row1 = ttk.Frame(card, style="Card.TFrame")
        row1.pack(fill="x", padx=16, pady=(14, 6))
        ttk.Label(row1, text="学生作业文件夹", style="Card.TLabel").pack(anchor="w")
        ttk.Label(
            row1,
            text="选「一人一个子文件夹」的总目录",
            style="Muted.TLabel",
        ).pack(anchor="w")
        r1b = ttk.Frame(card, style="Card.TFrame")
        r1b.pack(fill="x", padx=16, pady=(0, 10))
        self._entry(r1b, self.submissions).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(r1b, text="浏览…", command=self._browse_submissions).pack(
            side=RIGHT, padx=(8, 0)
        )

        # 输出
        row2 = ttk.Frame(card, style="Card.TFrame")
        row2.pack(fill="x", padx=16, pady=(6, 6))
        ttk.Label(row2, text="成绩表保存为", style="Card.TLabel").pack(anchor="w")
        r2b = ttk.Frame(card, style="Card.TFrame")
        r2b.pack(fill="x", padx=16, pady=(0, 10))
        self._entry(r2b, self.output).pack(side=LEFT, fill="x", expand=True)
        ttk.Button(r2b, text="浏览…", command=self._browse_output).pack(
            side=RIGHT, padx=(8, 0)
        )

        # 班级
        row3 = ttk.Frame(card, style="Card.TFrame")
        row3.pack(fill="x", padx=16, pady=(6, 4))
        ttk.Label(row3, text="成绩单「班级」列", style="Card.TLabel").pack(anchor="w")
        r3b = ttk.Frame(card, style="Card.TFrame")
        r3b.pack(fill="x", padx=16, pady=(0, 4))
        ttk.Radiobutton(
            r3b,
            text="自动（文件夹名里出现最多的班级）",
            variable=self.class_mode,
            value="auto",
            command=self._toggle_class,
        ).pack(anchor="w")
        ttk.Radiobutton(
            r3b,
            text="留空",
            variable=self.class_mode,
            value="empty",
            command=self._toggle_class,
        ).pack(anchor="w")
        r3c = ttk.Frame(card, style="Card.TFrame")
        r3c.pack(fill="x", padx=16, pady=(0, 8))
        ttk.Radiobutton(
            r3c,
            text="指定：",
            variable=self.class_mode,
            value="custom",
            command=self._toggle_class,
        ).pack(side=LEFT)
        self.class_entry = self._entry(r3c, self.class_custom, width=36)
        self.class_entry.pack(side=LEFT, padx=(4, 0))
        ttk.Label(
            card,
            text="任课班参考：" + "、".join(TEACHER_CLASSES[:4]) + "…",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=16, pady=(0, 8))

        row4 = ttk.Frame(card, style="Card.TFrame")
        row4.pack(fill="x", padx=16, pady=(0, 14))
        ttk.Checkbutton(
            row4,
            text="跳过答案自检（加快速度）",
            variable=self.skip_self_check,
        ).pack(side=LEFT)

        # 按钮区
        actions = ttk.Frame(self.root, style="TFrame")
        actions.pack(fill="x", padx=20, pady=10)
        self.btn_run = ttk.Button(
            actions, text="开始批改", style="Accent.TButton", command=self._start
        )
        self.btn_run.pack(side=LEFT)
        self.btn_open = ttk.Button(
            actions, text="打开成绩表", command=self._open_excel, state=DISABLED
        )
        self.btn_open.pack(side=LEFT, padx=(10, 0))
        self.btn_folder = ttk.Button(
            actions, text="打开输出目录", command=self._open_outdir, state=DISABLED
        )
        self.btn_folder.pack(side=LEFT, padx=(10, 0))
        self.status = ttk.Label(actions, text="就绪", style="Sub.TLabel")
        self.status.pack(side=RIGHT)

        # 日志
        log_frame = ttk.Frame(self.root, style="Card.TFrame")
        log_frame.pack(fill=BOTH, expand=True, padx=20, pady=(0, 16))
        ttk.Label(log_frame, text="运行日志", style="Card.TLabel").pack(
            anchor="w", padx=12, pady=(10, 4)
        )
        self.log = ScrolledText(
            log_frame,
            height=16,
            wrap=WORD,
            font=self.font_mono,
            bg=COLORS["input_bg"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            borderwidth=0,
        )
        self.log.pack(fill=BOTH, expand=True, padx=12, pady=(0, 12))
        self._log(f"{APP_TITLE}\n作者：{AUTHOR}\n程序目录：{app_dir()}\n")
        self._toggle_class()

    def _entry(self, parent, var: StringVar, width: int | None = None):
        e = ttk.Entry(parent, textvariable=var, font=self.font_ui)
        if width:
            e.configure(width=width)
        return e

    def _toggle_class(self):
        state = NORMAL if self.class_mode.get() == "custom" else DISABLED
        self.class_entry.configure(state=state)

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

    def _set_status(self, text: str):
        self.status.configure(text=text)

    def _start(self):
        if self.busy:
            return
        sub = Path(self.submissions.get().strip().strip('"'))
        out = Path(self.output.get().strip().strip('"'))
        if not sub.exists() or not sub.is_dir():
            messagebox.showwarning("提示", "请先选择有效的学生作业文件夹。")
            return
        if not out.parent.exists():
            messagebox.showwarning("提示", "成绩表保存目录不存在。")
            return

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
        self._set_status("批改中…")
        self._log("\n" + "=" * 40)
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
                    # self_check 用 print，这里直接调并捕获异常
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
                            f"批改中 {i}/{total}"
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
                err = f"{e}\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._done_err(err))

        threading.Thread(target=work, daemon=True).start()

    def _done_ok(self, out: Path, summary: str):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self.btn_open.configure(state=NORMAL)
        self.btn_folder.configure(state=NORMAL)
        self.last_output = out
        self._log(summary)
        self._set_status("完成")
        messagebox.showinfo("批改完成", summary)

    def _done_err(self, err: str):
        self.busy = False
        self.btn_run.configure(state=NORMAL)
        self._log("出错了：\n" + err)
        self._set_status("出错")
        messagebox.showerror("批改失败", err[:800])

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
    # Windows 高 DPI
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    GradeApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()
