# -*- coding: utf-8 -*-
"""
期中考试 Word 操作题自动批改
依据：评分标准.md
作者：Fantasy
"""
from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field, asdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from pypinyin import lazy_pinyin


def app_dir() -> Path:
    """exe / 脚本所在目录（用户放试卷、提交文件夹的地方）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    """打包进 exe 的资源目录；开发时等于项目根目录。"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def resolve_answer(name: str) -> Path:
    """优先用 exe 旁的「答案」文件夹，否则用内置答案。"""
    external = app_dir() / "答案" / name
    if external.exists():
        return external
    bundled = resource_dir() / "答案" / name
    return bundled


ROOT = app_dir()
ANSWER_TYPING = resolve_answer("打字测试.docx")
ANSWER_RESUME = resolve_answer("简历表.docx")
DEFAULT_OUTPUT = ROOT / "成绩表.xlsx"
# 成绩单「班级」列默认留空；可用 --class-name 或运行时输入
DEFAULT_CLASS_NAME = ""
AUTHOR = "Fantasy"
APP_TITLE = "Word 操作题自动批改"
# 列举文件夹时忽略这些名字
_SKIP_DIR_NAMES = {
    "答案",
    "assets",
    "dist",
    "build",
    "release",
    "local_sensitive",
    "__pycache__",
    ".git",
    ".idea",
    ".vscode",
}

# 试卷分值
SCORE_TYPING_COMPLETE = 20.0
SCORE_TYPING_ACCURACY = 20.0
SCORE_RESUME_TOTAL = 50.0
SCORE_NAMING = 10.0

RESUME_KEYWORDS = [
    "姓名",
    "国籍",
    "照片",
    "目前所在地",
    "民族",
    "户口所在地",
    "出生日期",
    "婚姻状况",
    "身高",
    "求职意向及工作经历",
    "人才类型",
    "应聘岗位",
    "工作类型",
    "最低薪资",
    "个人工作经历",
    "教育背景",
    "毕业学校",
    "最高学历",
    "毕业日期",
    "所学专业",
    "自我评价及爱好",
]

SECTION_LABELS = ("求职意向及工作经历", "教育背景", "自我评价及爱好")
VERTICAL_DIRS = {"tbRl", "btLr", "tbRlV", "lrTbV", "btLrV"}
GRAY_FILLS = {
    "D9D9D9",
    "C0C0C0",
    "BFBFBF",
    "A6A6A6",
    "808080",
    "F2F2F2",
    "E7E6E6",
    "D0D0D0",
    "B3B3B3",
    "AUTO",
}
# 任课班级（课表）：用于识别/归一化文件夹里的班级写法
TEACHER_CLASSES = [
    "26级美容保健1班",
    "26级美容保健2班",
    "26级电子商务1班",
    "26级电子商务2班",
    "25计算机网络应用1班",
    "26级计算机网络应用1班",
]
# 简写 → 标准班名（文件夹常写「26电商1班」）
CLASS_ALIASES = {
    "26电商1班": "26级电子商务1班",
    "26电商2班": "26级电子商务2班",
    "26电子商务1班": "26级电子商务1班",
    "26电子商务2班": "26级电子商务2班",
    "26美容保健1班": "26级美容保健1班",
    "26美容保健2班": "26级美容保健2班",
    "26美容1班": "26级美容保健1班",
    "26美容2班": "26级美容保健2班",
    "26级美容1班": "26级美容保健1班",
    "26级美容2班": "26级美容保健2班",
    "25计算机网络1班": "25计算机网络应用1班",
    "26计算机网络应用1班": "26级计算机网络应用1班",
    "26级计算机网络1班": "26级计算机网络应用1班",
}

# 文件夹里像「26美容保健1班」「26级电子商务1班」这类班级片段
CLASS_CHUNK = re.compile(r"\d+级?[\u4e00-\u9fff]+?\d*班")
CLASS_HINT = re.compile(
    r"(电商|电子商务|计算机|美容|网络应用|\d+级?[\u4e00-\u9fff]+?\d*班|班)"
)
CN_NAME = re.compile(r"[\u4e00-\u9fff]{2,4}")
IP_PREFIX = re.compile(r"^\d{1,3}(?:\.\d{1,3}){3}_?")
SEAT_PREFIX = re.compile(r"^[A-Za-z]\d{1,3}_?")
# 从整段里误提取时要丢掉的专业/班级碎片
_NAME_BAD = {
    "电商",
    "电子商务",
    "计算机",
    "级电",
    "子商务",
    "美容",
    "美容保健",
    "级美容",
    "保健",
    "网络",
    "网络应用",
    "计算机网",
    "算机网",
}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def round_half(x: float, nd: int = 1) -> float:
    return round(float(x) + 1e-9, nd)


def _ps_quote(s: str) -> str:
    return "'" + str(s).replace("'", "''") + "'"


def _run_powershell(script: str, timeout: int = 180) -> bool:
    try:
        r = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding="utf-8",
            errors="replace",
        )
        return r.returncode == 0
    except Exception:
        return False


def convert_doc_to_docx(src: Path) -> Optional[Path]:
    """
    把旧版 .doc 转成临时 .docx，供 python-docx 读取。
    优先 Microsoft Word COM，其次 WPS。
    """
    src = src.resolve()
    if not src.exists() or src.suffix.lower() != ".doc":
        return None

    cache_dir = Path(tempfile.gettempdir()) / "OfficeGrade_doc_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    digest = hashlib.md5(str(src).encode("utf-8")).hexdigest()[:12]
    out = cache_dir / f"{src.stem}_{digest}.docx"
    if out.exists() and out.stat().st_mtime >= src.stat().st_mtime:
        return out

    src_q, out_q = _ps_quote(str(src)), _ps_quote(str(out))

    # Word：SaveAs 格式 16 = wdFormatXMLDocument (.docx)
    word_script = f"""
$ErrorActionPreference = 'Stop'
$app = $null; $doc = $null
try {{
  $app = New-Object -ComObject Word.Application
  $app.Visible = $false
  $app.DisplayAlerts = 0
  $src = {src_q}
  $out = {out_q}
  $doc = $app.Documents.Open($src, $false, $true)
  if (Test-Path -LiteralPath $out) {{ Remove-Item -LiteralPath $out -Force }}
  $fmt = 16
  $null = $doc.SaveAs([ref]$out, [ref]$fmt)
}} finally {{
  if ($doc -ne $null) {{
    $c = $false
    $doc.Close([ref]$c) | Out-Null
  }}
  if ($app -ne $null) {{ $app.Quit() | Out-Null }}
}}
"""
    if _run_powershell(word_script) and out.exists():
        return out

    # WPS 兼容
    wps_script = f"""
$ErrorActionPreference = 'Stop'
$app = $null; $doc = $null
try {{
  try {{ $app = New-Object -ComObject Kwps.Application }} catch {{ $app = New-Object -ComObject wps.Application }}
  $app.Visible = $false
  $src = {src_q}
  $out = {out_q}
  $doc = $app.Documents.Open($src)
  if (Test-Path -LiteralPath $out) {{ Remove-Item -LiteralPath $out -Force }}
  $null = $doc.SaveAs($out, 12)
}} finally {{
  if ($doc -ne $null) {{ $doc.Close() | Out-Null }}
  if ($app -ne $null) {{ $app.Quit() | Out-Null }}
}}
"""
    if _run_powershell(wps_script) and out.exists():
        return out
    return None


def ensure_docx(path: Path) -> Optional[Path]:
    """返回可被 python-docx 打开的 .docx 路径；.doc 会先转换。"""
    if path is None or not path.exists():
        return None
    if path.name.startswith("~$"):
        return None
    suf = path.suffix.lower()
    if suf == ".docx":
        return path
    if suf == ".doc":
        return convert_doc_to_docx(path)
    return None


def safe_open_docx(path: Path) -> Optional[Document]:
    """打开 .docx；若是 .doc 则先转成临时 docx 再打开。"""
    readable = ensure_docx(path)
    if readable is None:
        return None
    try:
        return Document(str(readable))
    except Exception:
        return None


def is_word_submission(path: Path) -> bool:
    return path.suffix.lower() in {".doc", ".docx"} and not path.name.startswith("~$")


def para_full_text(doc: Document) -> str:
    parts = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)
    # 表格内文字也拼上（少数学生把内容放进表格）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    parts.append(t)
    return "\n".join(parts)


def normalize_typing(text: str) -> str:
    text = text.replace("\r", "").replace("\u3000", "").replace(" ", "")
    text = re.sub(r"\n+", "", text)
    return text.strip()


def split_sentences(text: str) -> list[str]:
    text = normalize_typing(text)
    if not text:
        return []
    parts = re.split(r"(?<=[。！？])", text)
    return [p for p in parts if p]


def edit_distance(a: str, b: str) -> int:
    """Levenshtein distance (char-level)."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    # 两行 DP 省内存
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost))
        prev = cur
    return prev[-1]


def _clip(s: str, n: int = 24) -> str:
    s = s.replace("\n", "")
    if len(s) <= n:
        return s
    return s[:n] + "…"


def _around(text: str, start: int, end: int, radius: int = 10) -> str:
    """取出差异处前后文，中间用【】标出问题字，方便定位。"""
    left = text[max(0, start - radius) : start]
    mid = text[start:end]
    right = text[end : end + radius]
    prefix = "…" if start - radius > 0 else ""
    suffix = "…" if end + radius < len(text) else ""
    if mid:
        return f"{prefix}{left}【{mid}】{right}{suffix}"
    # 插入点：在答案侧没有 mid，用 ∧ 标出位置
    return f"{prefix}{left}【∧】{right}{suffix}"


def explain_char_diff(answer: str, student: str, max_items: int = 20) -> tuple[int, list[str]]:
    """
    用 SequenceMatcher 列出字级差异，返回 (差异字数约值, 白话条目列表)。
    每条都带前后文，标明错在哪一段。
    """
    dist = edit_distance(answer, student)
    if dist == 0:
        return 0, []

    sm = SequenceMatcher(None, answer, student, autojunk=False)
    items: list[str] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        a_frag = answer[i1:i2]
        b_frag = student[j1:j2]
        a_ctx = _around(answer, i1, i2)
        b_ctx = _around(student, j1, j2)
        if tag == "replace":
            items.append(
                f"写错了：答案片段「{a_ctx}」→ 你写成「{b_ctx}」"
                f"（错字：答案「{_clip(a_frag)}」你写成「{_clip(b_frag)}」）"
            )
        elif tag == "delete":
            items.append(
                f"少打了：答案里「{a_ctx}」这一段有「{_clip(a_frag)}」，"
                f"你这边对应位置附近是「{b_ctx}」，漏了这些字"
            )
        elif tag == "insert":
            items.append(
                f"多打了：你在「{b_ctx}」多打了「{_clip(b_frag)}」；"
                f"答案对应位置附近是「{a_ctx}」，本来没有这些字"
            )
        if len(items) >= max_items:
            items.append(f"……还有更多差异（共差约{dist}字，此处只列前{max_items}处）")
            break
    if not items:
        items.append(f"正文与答案不一致，共差约{dist}字（无法逐条拆开时）")
    return dist, items


def run_east_asia_font(run) -> Optional[str]:
    try:
        rPr = run._element.rPr
        if rPr is None or rPr.rFonts is None:
            return None
        return rPr.rFonts.get(qn("w:eastAsia")) or run.font.name
    except Exception:
        return run.font.name


def is_songti(name: Optional[str]) -> bool:
    """未显式设置字体时，中文 Word 默认多为宋体，从宽视为合格。"""
    if not name:
        return True
    n = name.replace(" ", "")
    if n in {"宋体", "SimSun", "NSimSun", "新宋体", "STSong", "方正宋体"}:
        return True
    # 主题字体如 +EastAsia 等从宽
    if "宋" in n:
        return True
    return False


def pt_of(run) -> Optional[float]:
    if run.font.size is None:
        return None
    return run.font.size.pt


def cell_shading(cell) -> Optional[str]:
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    return fill.upper() if fill else None


def cell_text_direction(cell) -> Optional[str]:
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return None
    td = tcPr.find(qn("w:textDirection"))
    if td is None:
        return None
    return td.get(qn("w:val"))


def cell_valign(cell) -> Optional[str]:
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return None
    va = tcPr.find(qn("w:vAlign"))
    if va is None:
        return None
    return va.get(qn("w:val"))


def cell_border_vals(cell) -> dict[str, str]:
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return {}
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        return {}
    out = {}
    for edge in ("top", "left", "bottom", "right"):
        el = borders.find(qn(f"w:{edge}"))
        if el is not None:
            out[edge] = el.get(qn("w:val")) or ""
    return out


def table_border_vals(table) -> dict[str, str]:
    tblPr = table._tbl.tblPr
    if tblPr is None:
        return {}
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        return {}
    out = {}
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is not None:
            out[edge] = el.get(qn("w:val")) or ""
    return out


def unique_cells(table):
    """按底层 w:tc 元素去重（勿用 id()，以免代理对象回收导致误判）。"""
    seen: set = set()
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            tc = cell._tc
            if tc in seen:
                continue
            seen.add(tc)
            yield ri, ci, cell


def table_all_text(doc: Document) -> str:
    parts: list[str] = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text)
    for table in doc.tables:
        # 完整性检测不去重，避免漏读合并单元格展示文本
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return "\n".join(parts)


def iter_cell_runs(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            if r.text is None:
                continue
            yield p, r


# ---------------------------------------------------------------------------
# student discovery
# ---------------------------------------------------------------------------

@dataclass
class StudentWork:
    folder_name: str
    folder_path: Path
    class_name: str = ""
    student_name: str = ""
    typing_path: Optional[Path] = None
    resume_path: Optional[Path] = None
    notes: list[str] = field(default_factory=list)


def parse_folder_identity(folder_name: str) -> tuple[str, str]:
    """
    从文件夹名提取班级、姓名。

    可靠规则：找到「…班」后，后面紧跟/空格分隔的中文才是姓名。
    例：L14_26美容保健1班易楚瑶 → 班级=26美容保健1班，姓名=易楚瑶
    """
    raw = folder_name
    # 去掉 IP / 座位号前缀
    s = IP_PREFIX.sub("", raw)
    s = SEAT_PREFIX.sub("", s)
    s = s.replace("_", " ").replace("-", " ").strip()

    class_name = ""
    student_name = ""

    # 优先：数字(+级)?专业名+班 + 可选姓名
    m = re.search(
        r"(\d+级?[\u4e00-\u9fff]+?\d*班)\s*([\u4e00-\u9fff]{2,4})?",
        s,
    )
    if m:
        class_name = normalize_class_name(m.group(1))
        if m.group(2):
            student_name = m.group(2)
        else:
            rest = s[m.end() :].strip()
            nm = CN_NAME.search(rest)
            if nm:
                student_name = nm.group(0)
        return class_name, student_name

    # 退路：没有「xx班」结构时，取末尾像姓名的 2~4 个汉字
    names = [
        n
        for n in CN_NAME.findall(s)
        if n not in _NAME_BAD and "班" not in n and not n.startswith("级")
    ]
    student_name = names[-1] if names else ""
    return class_name, student_name


def normalize_class_name(name: str) -> str:
    """把文件夹里的班级简写归一成课表标准班名。"""
    name = (name or "").strip()
    if not name:
        return ""
    if name in TEACHER_CLASSES:
        return name
    if name in CLASS_ALIASES:
        return CLASS_ALIASES[name]
    # 再试去掉空格
    compact = name.replace(" ", "")
    if compact in TEACHER_CLASSES:
        return compact
    if compact in CLASS_ALIASES:
        return CLASS_ALIASES[compact]
    return name


def is_exam_paper(path: Path, doc: Optional[Document] = None) -> bool:
    if "试卷" in path.stem:
        return True
    if doc is None:
        return False
    text = para_full_text(doc)[:200]
    return "期中考试试卷" in text or "绝密" in text


def classify_docx(path: Path) -> str:
    """返回 typing / resume / exam / unknown / invalid"""
    name = path.name
    if name.startswith("~$") or path.suffix.lower() == ".lnk":
        return "invalid"
    if path.suffix.lower() not in {".doc", ".docx"}:
        return "invalid"
    stem = path.stem
    # 双后缀
    if stem.endswith(".docx"):
        stem = stem[:-5]

    doc = safe_open_docx(path)
    if doc is None:
        # .doc 转换失败时仍按文件名粗分，便于备注提示
        if path.suffix.lower() == ".doc":
            if "打字" in stem:
                return "typing"
            if "简历" in stem:
                return "resume"
            return "unknown"
        return "invalid"
    if is_exam_paper(path, doc):
        return "exam"

    text = para_full_text(doc)
    norm = normalize_typing(text)
    # 打字题特征
    typing_hints = ("工匠精神", "大国工匠", "技工学校")
    resume_hints = ("个人简历表", "求职意向", "教育背景", "自我评价")

    name_typing = "打字" in stem or stem in {"打字测试"}
    name_resume = any(k in stem for k in ("简历", "个人简历"))

    score_t = sum(1 for h in typing_hints if h in norm) + (2 if name_typing else 0)
    score_r = sum(1 for h in resume_hints if h in text) + (2 if name_resume else 0)
    # 有表格更像简历
    if doc.tables:
        score_r += 2

    if score_t >= 2 and score_t > score_r:
        return "typing"
    if score_r >= 2 and score_r > score_t:
        return "resume"
    if name_typing:
        return "typing"
    if name_resume:
        return "resume"
    if score_t >= 1:
        return "typing"
    if score_r >= 1:
        return "resume"
    return "unknown"


def pick_best(files: list[Path], kind: str) -> Optional[Path]:
    if not files:
        return None
    if kind == "typing":
        preferred = ["打字测试.docx", "打字测试.doc", "打字测试"]
    else:
        preferred = [
            "简历表.docx",
            "简历表.doc",
            "个人简历表.docx",
            "个人简历表.doc",
            "简历表",
            "个人简历表",
        ]

    def rank(p: Path) -> tuple:
        stem = p.name
        base = p.stem
        if base.endswith(".docx"):
            base = base[:-5]
        for i, pref in enumerate(preferred):
            if stem == pref or base == pref.replace(".docx", "").replace(".doc", ""):
                # 同名时优先 .docx
                return (0, i, 0 if p.suffix.lower() == ".docx" else 1, len(stem))
        return (1, 0, 0 if p.suffix.lower() == ".docx" else 1, len(stem))

    return sorted(files, key=rank)[0]


def discover_students(submissions: Path) -> list[StudentWork]:
    students: list[StudentWork] = []

    # 子文件夹（跳过答案、打包与工程目录，避免误批）
    for d in sorted(submissions.iterdir()):
        if not d.is_dir():
            continue
        if d.name.startswith(".") or d.name in _SKIP_DIR_NAMES:
            continue
        sw = StudentWork(folder_name=d.name, folder_path=d)
        sw.class_name, sw.student_name = parse_folder_identity(d.name)

        typing_cands, resume_cands = [], []
        for f in d.rglob("*"):
            if not f.is_file():
                continue
            suf = f.suffix.lower()
            if suf not in {".doc", ".docx", ".lnk"}:
                continue
            kind = classify_docx(f)
            if kind == "typing":
                typing_cands.append(f)
                if suf == ".doc" and ensure_docx(f) is None:
                    sw.notes.append(
                        f"【提示】「{f.name}」是旧版.doc，转换失败（需本机安装 Word 或 WPS），内容可能批不了。"
                    )
            elif kind == "resume":
                resume_cands.append(f)
                if suf == ".doc" and ensure_docx(f) is None:
                    sw.notes.append(
                        f"【提示】「{f.name}」是旧版.doc，转换失败（需本机安装 Word 或 WPS），内容可能批不了。"
                    )
            elif kind == "exam":
                sw.notes.append(
                    f"【提示】「{f.name}」是试卷原件，已忽略，不作为作业批改。"
                )
            elif kind == "invalid":
                if suf == ".lnk":
                    sw.notes.append(
                        f"【提示】「{f.name}」只是快捷方式(.lnk)，不是真正的Word文件，"
                        "系统打不开里面的内容，等于没交这份作业。"
                    )
                elif f.name.startswith("~$"):
                    pass
                else:
                    sw.notes.append(f"【提示】「{f.name}」无法作为Word打开。")
            else:
                sw.notes.append(
                    f"【提示】「{f.name}」既不像打字题也不像简历表，未纳入批改。"
                )

        sw.typing_path = pick_best(typing_cands, "typing")
        sw.resume_path = pick_best(resume_cands, "resume")
        if len(typing_cands) > 1 and sw.typing_path:
            sw.notes.append(
                f"【提示】打字相关文件有多份，系统选了最像标准文件名的：「{sw.typing_path.name}」。"
            )
        if len(resume_cands) > 1 and sw.resume_path:
            sw.notes.append(
                f"【提示】简历相关文件有多份，系统选了最像标准文件名的：「{sw.resume_path.name}」。"
            )
        students.append(sw)

    # 根目录散落的单个 doc/docx：每人一个「伪文件夹」
    for f in sorted(submissions.iterdir()):
        if not f.is_file() or f.suffix.lower() not in {".doc", ".docx"}:
            continue
        if f.name.startswith("~$"):
            continue
        kind = classify_docx(f)
        cls, name = parse_folder_identity(f.stem)
        sw = StudentWork(
            folder_name=f"(根目录){f.stem}",
            folder_path=submissions,
            class_name=cls,
            student_name=name or f.stem,
        )
        sw.notes.append(
            "【提示】这个文件直接放在收取根目录，没有「班级+姓名」文件夹。"
        )
        if kind == "typing":
            sw.typing_path = f
        elif kind == "resume":
            sw.resume_path = f
        elif kind == "exam":
            sw.notes.append("【提示】根目录这个文件是试卷，不是作答。")
        else:
            sw.typing_path = f
            sw.notes.append(
                f"【提示】根目录「{f.name}」类型不清，先按打字题试批；请人工确认。"
            )
        students.append(sw)

    return students


# ---------------------------------------------------------------------------
# scoring: naming
# ---------------------------------------------------------------------------

def score_naming(sw: StudentWork) -> tuple[float, list[str]]:
    score = SCORE_NAMING
    reasons = []

    folder = sw.folder_name
    in_root = folder.startswith("(根目录)")
    folder_ok = False
    if not in_root:
        has_class = bool(CLASS_HINT.search(folder))
        has_name = bool(sw.student_name) or bool(
            CN_NAME.search(SEAT_PREFIX.sub("", IP_PREFIX.sub("", folder)))
        )
        folder_ok = has_class and has_name
    if not folder_ok:
        score -= 4
        if in_root:
            reasons.append(
                "【文件命名 −4分】试卷要求：在桌面建文件夹，名字=班级+姓名"
                f"（例：26电子商务1班张三）。你没有独立文件夹，文件直接扔在收取目录里"
                f"（{folder}），所以扣4分。"
            )
        else:
            reasons.append(
                "【文件命名 −4分】试卷要求文件夹叫「班级+姓名」。"
                f"你的文件夹是「{folder}」，看不出完整班级名（要有「电商/电子商务/…班」这类字样）"
                "或姓名不清晰，所以扣4分。"
            )

    def typing_name_ok(p: Optional[Path]) -> bool:
        if p is None:
            return False
        # .doc / .docx 都认；双后缀 .docx.docx 不算
        return p.name.strip() in {"打字测试.docx", "打字测试.doc"}

    def resume_name_ok(p: Optional[Path]) -> bool:
        if p is None:
            return False
        return p.name.strip() in {
            "简历表.docx",
            "简历表.doc",
            "个人简历表.docx",
            "个人简历表.doc",
        }

    if not typing_name_ok(sw.typing_path):
        score -= 3
        actual = sw.typing_path.name if sw.typing_path else "（根本没交打字文件）"
        reasons.append(
            "【文件命名 −3分】打字文档的文件名必须是「打字测试.docx」或「打字测试.doc」"
            f"（多一个字、少一个字、双后缀.docx.docx都不行）。你实际是「{actual}」，扣3分。"
        )
    if not resume_name_ok(sw.resume_path):
        score -= 3
        actual = sw.resume_path.name if sw.resume_path else "（根本没交简历文件）"
        reasons.append(
            "【文件命名 −3分】简历文档文件名必须是「简历表.docx/.doc」或「个人简历表.docx/.doc」。"
            f"你实际是「{actual}」，扣3分。"
        )

    return max(0.0, round_half(score)), reasons


# ---------------------------------------------------------------------------
# scoring: typing
# ---------------------------------------------------------------------------

def score_typing(path: Optional[Path], answer_text: str) -> tuple[float, float, list[str]]:
    """返回 (完整度, 准确率, 白话扣分说明)"""
    notes: list[str] = []
    if path is None:
        return (
            0.0,
            0.0,
            [
                "【打字测试 0分】没找到打字文档。完整度20分+准确率20分全部没有。"
                "常见原因：没交、只交了快捷方式.lnk、或文件坏了打不开。"
            ],
        )
    doc = safe_open_docx(path)
    if doc is None:
        return (
            0.0,
            0.0,
            [f"【打字测试 0分】文件「{path.name}」打不开（不是正常docx或已损坏）。"],
        )
    if is_exam_paper(path, doc):
        return (
            0.0,
            0.0,
            [
                f"【打字测试 0分】你交的「{path.name}」看起来是试卷原件，不是自己录入的打字作业。"
            ],
        )

    student = normalize_typing(para_full_text(doc))
    answer_n = normalize_typing(answer_text)
    if not student:
        return 0.0, 0.0, [f"【打字测试 0分】「{path.name}」里几乎没有正文，像空文档。"]

    # 完整度：按句
    sentences = split_sentences(answer_n)
    n = max(len(sentences), 1)
    per = SCORE_TYPING_COMPLETE / n
    missing_sents: list[str] = []
    for sent in sentences:
        if sent in student:
            continue
        best = 0.0
        L = len(sent)
        if L == 0:
            continue
        if len(student) <= L:
            best = SequenceMatcher(None, sent, student).ratio()
        else:
            step = max(1, L // 4)
            for i in range(0, len(student) - L + 1, step):
                frag = student[i : i + L]
                best = max(best, SequenceMatcher(None, sent, frag).ratio())
                if best > 0.85:
                    break
        if best < 0.55:
            missing_sents.append(sent)

    missing = len(missing_sents)
    complete = max(0.0, SCORE_TYPING_COMPLETE - missing * per)
    complete = round_half(complete)

    # 准确率 + 逐条差异
    dist, diff_items = explain_char_diff(answer_n, student)
    accuracy = round_half(max(0.0, SCORE_TYPING_ACCURACY - dist))

    if missing:
        lost = round_half(SCORE_TYPING_COMPLETE - complete)
        lines = [
            f"【打字完整度 满分20，得{complete}，扣{lost}分】"
            f"标准答案拆成{n}句，漏一整句大约扣{per:.1f}分。"
            f"系统判断你漏了{missing}整句："
        ]
        for i, s in enumerate(missing_sents, 1):
            lines.append(f"  {i}. 漏句：「{_clip(s, 40)}」")
        notes.append("\n".join(lines))

    if dist:
        lost = round_half(SCORE_TYPING_ACCURACY - accuracy)
        lines = [
            f"【打字准确率 满分20，得{accuracy}，扣{lost}分】"
            f"规则：和标准答案比，每多1字/少1字/错1字（含标点）扣1分，扣完为止。"
            f"你这边一共差{dist}字。具体是这些地方不对："
        ]
        for i, item in enumerate(diff_items, 1):
            lines.append(f"  {i}. {item}")
        notes.append("\n".join(lines))
    elif not missing:
        notes.append("【打字测试】完整度、准确率均与答案一致，本大题内容分满分。")

    return complete, accuracy, notes


# ---------------------------------------------------------------------------
# scoring: resume
# ---------------------------------------------------------------------------

@dataclass
class ResumeScores:
    completeness: float = 0.0
    title: float = 0.0
    font_bold: float = 0.0
    photo: float = 0.0
    double_border: float = 0.0
    gray: float = 0.0
    notes: list[str] = field(default_factory=list)

    @property
    def total(self) -> float:
        return round_half(
            self.completeness
            + self.title
            + self.font_bold
            + self.photo
            + self.double_border
            + self.gray
        )


def find_title_paragraph(doc: Document):
    for p in doc.paragraphs:
        if "个人简历表" in p.text.replace(" ", ""):
            return p
    # 有的写在首格
    for table in doc.tables:
        for _, _, cell in unique_cells(table):
            if cell.text.strip().replace(" ", "") == "个人简历表":
                # 伪造一个简易对象：用 cell 首段
                if cell.paragraphs:
                    return cell.paragraphs[0]
    return None


def score_title(doc: Document) -> tuple[float, list[str]]:
    notes = []
    p = find_title_paragraph(doc)
    if p is None:
        return 0.0, [
            "【简历·标题样式 满分5，得0，扣5分】试卷要求表上方写「个人简历表」，"
            "并且：宋体、二号、加粗、居中。你的文档里找不到这行标题，所以这项0分。"
        ]
    checks = {"宋体": False, "二号(22磅)": False, "加粗": False, "居中": False}
    for r in p.runs:
        if not r.text.strip():
            continue
        if is_songti(run_east_asia_font(r)) or is_songti(r.font.name):
            checks["宋体"] = True
        sz = pt_of(r)
        if sz is not None and abs(sz - 22.0) < 0.6:
            checks["二号(22磅)"] = True
        if r.bold:
            checks["加粗"] = True
    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        checks["居中"] = True
    if not checks["二号(22磅)"]:
        try:
            if p.style and p.style.font and p.style.font.size:
                if abs(p.style.font.size.pt - 22.0) < 0.6:
                    checks["二号(22磅)"] = True
        except Exception:
            pass

    each = 5.0 / 4
    got = sum(1 for v in checks.values() if v) * each
    got = round_half(got)
    miss = [k for k, v in checks.items() if not v]
    if miss:
        lost = round_half(5.0 - got)
        notes.append(
            f"【简历·标题样式 满分5，得{got}，扣{lost}分】"
            f"「个人简历表」这四个要求里，你还缺：{'、'.join(miss)}。"
            f"四项各约1.25分，缺一项就扣一项的分。"
        )
    return got, notes


def score_completeness(doc: Document) -> tuple[float, list[str]]:
    if not doc.tables:
        return 0.0, [
            "【简历·表格完整度 满分10，得0，扣10分】文档里没有表格。"
            "试卷要求做成一张简历表（有行列的那种），你没做出来。"
        ]
    text = table_all_text(doc)
    hit_list = [kw for kw in RESUME_KEYWORDS if kw in text]
    miss_list = [kw for kw in RESUME_KEYWORDS if kw not in text]
    hit = len(hit_list)
    ratio = hit / len(RESUME_KEYWORDS)
    score = round_half(min(10.0, 10.0 * min(1.0, ratio / 0.9)))
    notes = []
    if ratio < 0.9:
        lost = round_half(10.0 - score)
        show_miss = "、".join(miss_list[:12])
        if len(miss_list) > 12:
            show_miss += f"等共{len(miss_list)}项"
        notes.append(
            f"【简历·表格完整度 满分10，得{score}，扣{lost}分】"
            f"要对照图例把标签文字做齐（姓名、国籍、照片、求职意向…）。"
            f"标准关键词共{len(RESUME_KEYWORDS)}个，你命中{hit}个（≥90%才满分）。"
            f"你缺少的标签包括：{show_miss}。"
        )
    return score, notes


def score_font_and_bold(doc: Document) -> tuple[float, list[str]]:
    notes = []
    if not doc.tables:
        return 0.0, [
            "【简历·字体与加粗 满分5，得0，扣5分】没有表格，无法检查表内字体和加粗。"
        ]

    total_runs = 0
    ok_font = 0
    ok_size = 0
    ok_align = 0
    bold_hits = {lab: False for lab in SECTION_LABELS}

    for table in doc.tables:
        for _, _, cell in unique_cells(table):
            cell_txt = cell.text.replace(" ", "").replace("\n", "")
            for lab in SECTION_LABELS:
                if lab in cell_txt and not bold_hits[lab]:
                    # 单元格含标签且任一 run 加粗即认可（文字可能被拆成多个 run）
                    for _, r in iter_cell_runs(cell):
                        if r.bold is True:
                            bold_hits[lab] = True
                            break
            for p, r in iter_cell_runs(cell):
                if not (r.text or "").strip():
                    continue
                # 跳过标题级大字
                sz = pt_of(r)
                if sz is not None and sz >= 18:
                    continue
                total_runs += 1
                font = run_east_asia_font(r) or r.font.name
                if is_songti(font):
                    ok_font += 1
                if sz is None or abs(sz - 10.5) < 0.6:
                    # 未显式字号时从宽算五号（Word 常继承正文）
                    ok_size += 1
                if p.alignment == WD_ALIGN_PARAGRAPH.CENTER or p.alignment is None:
                    # None 时很多答案也靠表格默认；稍从宽：居中或未设都算
                    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        ok_align += 1
                    else:
                        # 未设置：记一半权重 —— 这里简化：不算对齐成功
                        pass
                if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    pass

    # 重新统计对齐（按段落）
    para_n = 0
    para_c = 0
    for table in doc.tables:
        for _, _, cell in unique_cells(table):
            for p in cell.paragraphs:
                if not p.text.strip():
                    continue
                para_n += 1
                if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    para_c += 1

    if total_runs == 0:
        return 0.0, [
            "【简历·字体与加粗 满分5，得0，扣5分】表格里几乎没有可检查的文字。"
        ]

    font_ratio = ok_font / total_runs
    size_ratio = ok_size / total_runs
    align_ratio = (para_c / para_n) if para_n else 0

    fmt = 0.0
    fmt_miss = []
    if font_ratio >= 0.6:
        fmt += 1.0
    else:
        fmt += font_ratio
        fmt_miss.append(f"宋体不够（达标约{font_ratio:.0%}，要≥60%）")
    if size_ratio >= 0.6:
        fmt += 1.0
    else:
        fmt += size_ratio
        fmt_miss.append(f"五号字不够（达标约{size_ratio:.0%}，要≥60%）")
    if align_ratio >= 0.6:
        fmt += 1.0
    else:
        fmt += align_ratio
        fmt_miss.append(f"水平居中不够（达标约{align_ratio:.0%}，要≥60%）")

    bold_score = 0.0
    each = 2.0 / 3
    for lab, hit in bold_hits.items():
        if hit:
            bold_score += each

    if bold_score < 2.0:
        for table in doc.tables:
            for _, _, cell in unique_cells(table):
                ct = cell.text.replace(" ", "").replace("\n", "")
                for lab in SECTION_LABELS:
                    if bold_hits[lab]:
                        continue
                    if lab not in ct:
                        continue
                    bolds = [
                        r.bold
                        for _, r in iter_cell_runs(cell)
                        if (r.text or "").strip()
                    ]
                    if bolds and any(b is True for b in bolds):
                        bold_hits[lab] = True
        bold_score = sum(each for lab in SECTION_LABELS if bold_hits[lab])

    bold_miss = [lab for lab, hit in bold_hits.items() if not hit]
    got = round_half(min(3.0, fmt) + min(2.0, bold_score))
    if fmt_miss or bold_miss:
        lost = round_half(5.0 - got)
        parts = [
            f"【简历·字体与加粗 满分5，得{got}，扣{lost}分】"
            "这项=表内宋体五号居中（3分）+三处标题加粗（2分）。"
        ]
        if fmt_miss:
            parts.append("格式问题：" + "；".join(fmt_miss) + "。")
        if bold_miss:
            parts.append(
                "这三处必须加粗：「求职意向及工作经历」「教育背景」「自我评价及爱好」。"
                f"你没加粗的是：{'、'.join(bold_miss)}。"
            )
        notes.append("".join(parts))

    return got, notes


def score_photo(doc: Document) -> tuple[float, list[str]]:
    notes = []
    if not doc.tables:
        return 0.0, [
            "【简历·照片纵向 满分10，得0，扣10分】没有表格，谈不上「照片」单元格。"
        ]
    photo_cell = None
    for table in doc.tables:
        for _, _, cell in unique_cells(table):
            t = cell.text.replace(" ", "").replace("\n", "")
            if t == "照片" or (t.startswith("照片") and len(t) <= 4):
                photo_cell = cell
                break
        if photo_cell:
            break
    if photo_cell is None:
        return 0.0, [
            "【简历·照片纵向 满分10，得0，扣10分】"
            "试卷要求：合并右上角单元格，里面写「照片」，文字方向改成纵向（竖着排），"
            "并且单元格居中。你的表里找不到「照片」这两个字，所以这项0分。"
        ]

    score = 0.0
    detail_miss = []
    tdir = cell_text_direction(photo_cell)
    if tdir in VERTICAL_DIRS:
        score += 6.0
    else:
        detail_miss.append("文字方向不是纵向/竖排（应在单元格里设「文字方向→纵向」），这项6分没拿到")

    horiz = any(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in photo_cell.paragraphs)
    vert = cell_valign(photo_cell) == "center"
    if horiz and vert:
        score += 4.0
    else:
        if horiz:
            score += 2.0
        else:
            detail_miss.append("「照片」没有水平居中（约2分）")
        if vert:
            score += 2.0
        else:
            detail_miss.append("「照片」单元格没有垂直居中（约2分）")

    score = round_half(score)
    if detail_miss:
        lost = round_half(10.0 - score)
        notes.append(
            f"【简历·照片纵向 满分10，得{score}，扣{lost}分】"
            + "；".join(detail_miss)
            + "。"
        )
    return score, notes


def find_section_rows(table) -> dict[str, list]:
    """label -> list of cells in that logical row (unique)."""
    found = {lab: [] for lab in SECTION_LABELS}
    # 先找标签所在行号
    label_row = {}
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            t = cell.text.replace(" ", "").replace("\n", "")
            for lab in SECTION_LABELS:
                if lab in t and lab not in label_row:
                    label_row[lab] = ri
    for lab, ri in label_row.items():
        if ri >= len(table.rows):
            continue
        seen: set = set()
        for cell in table.rows[ri].cells:
            tc = cell._tc
            if tc in seen:
                continue
            seen.add(tc)
            found[lab].append(cell)
    return found


def row_has_double(cells: list) -> bool:
    if not cells:
        return False
    double_edges = 0
    total_edges = 0
    for cell in cells:
        borders = cell_border_vals(cell)
        for v in borders.values():
            total_edges += 1
            if v == "double":
                double_edges += 1
    if total_edges == 0:
        # 尝试表级边框（少见）
        return False
    return double_edges / total_edges >= 0.4 or double_edges >= 2


def row_has_gray(cells: list) -> bool:
    if not cells:
        return False
    gray = 0
    for cell in cells:
        fill = cell_shading(cell)
        if fill and (fill in GRAY_FILLS or _is_gray_hex(fill)):
            gray += 1
    return gray >= max(1, len(cells) // 2) or gray >= 1


def _is_gray_hex(fill: str) -> bool:
    if not fill or fill.upper() == "AUTO":
        return False
    if len(fill) != 6:
        return False
    try:
        r, g, b = int(fill[0:2], 16), int(fill[2:4], 16), int(fill[4:6], 16)
    except ValueError:
        return False
    # 灰色：三通道接近且非纯白/纯黑
    if max(r, g, b) - min(r, g, b) > 15:
        return False
    return 40 <= r <= 230


def score_double_and_gray(doc: Document) -> tuple[float, float, list[str]]:
    notes = []
    if not doc.tables:
        return (
            0.0,
            0.0,
            [
                "【简历·双线/灰底 各满分10，得0】没有表格，双线边框和灰色底纹都没法给分。"
            ],
        )
    table = doc.tables[0]
    sections = find_section_rows(table)
    each_d = 10.0 / 3
    each_g = 10.0 / 3
    d_score = 0.0
    g_score = 0.0
    d_miss = []
    g_miss = []
    for lab in SECTION_LABELS:
        cells = sections[lab]
        if not cells:
            d_miss.append(f"找不到「{lab}」这一行")
            g_miss.append(f"找不到「{lab}」这一行")
            continue
        if row_has_double(cells):
            d_score += each_d
        else:
            d_miss.append(lab)
        if row_has_gray(cells):
            g_score += each_g
        else:
            g_miss.append(lab)
    d_score = round_half(d_score)
    g_score = round_half(g_score)
    if d_miss:
        notes.append(
            f"【简历·双线边框 满分10，得{d_score}，扣{round_half(10-d_score)}分】"
            "试卷要求第5/9/13行（求职意向、教育背景、自我评价）边框改成0.5磅双线。"
            f"三行大约各3.3分。你没做好的是：{'、'.join(d_miss)}。"
        )
    if g_miss:
        notes.append(
            f"【简历·灰色底纹 满分10，得{g_score}，扣{round_half(10-g_score)}分】"
            "同上三行要填充灰色底纹（浅灰即可）。"
            f"你没做好的是：{'、'.join(g_miss)}。"
        )
    return d_score, g_score, notes


def score_resume(path: Optional[Path]) -> ResumeScores:
    rs = ResumeScores()
    if path is None:
        rs.notes.append(
            "【简历表 整题0分】没找到简历文档（满分50分全无）。"
            "常见原因：没交、只交了.lnk快捷方式、或文件名完全对不上。"
        )
        return rs
    doc = safe_open_docx(path)
    if doc is None:
        rs.notes.append(
            f"【简历表 整题0分】文件「{path.name}」打不开，无法批改。"
        )
        return rs
    if is_exam_paper(path, doc):
        rs.notes.append(
            f"【简历表 整题0分】「{path.name}」是试卷原件，不是你做的简历表。"
        )
        return rs

    c, n = score_completeness(doc)
    rs.completeness, rs.notes = c, rs.notes + n
    t, n = score_title(doc)
    rs.title, rs.notes = t, rs.notes + n
    f, n = score_font_and_bold(doc)
    rs.font_bold, rs.notes = f, rs.notes + n
    p, n = score_photo(doc)
    rs.photo, rs.notes = p, rs.notes + n
    d, g, n = score_double_and_gray(doc)
    rs.double_border, rs.gray, rs.notes = d, g, rs.notes + n
    # 简历满分且无扣分说明时，补一句达标说明（避免备注只写打字）
    if rs.total >= 49.5 and not rs.notes:
        rs.notes.append(
            "【简历表】完整度、标题样式、字体加粗、照片纵向、双线边框、灰色底纹"
            "均达标，本大题满分50分。"
        )
    return rs


# ---------------------------------------------------------------------------
# excel / precheck / output helpers
# ---------------------------------------------------------------------------

@dataclass
class GradeRow:
    class_name: str
    student_name: str
    folder: str
    typing_file: str
    resume_file: str
    typing_complete: float
    typing_accuracy: float
    typing_total: float
    resume_complete: float
    resume_title: float
    resume_font: float
    resume_photo: float
    resume_double: float
    resume_gray: float
    resume_total: float
    naming: float
    total: float
    remarks: str
    # 复核 / 教务对接扩展（默认不影响旧逻辑）
    student_id: str = ""
    folder_path: str = ""
    typing_fullpath: str = ""
    resume_fullpath: str = ""
    final_total: Optional[float] = None
    adjust_note: str = ""

    def effective_total(self) -> float:
        if self.final_total is None:
            return self.total
        return self.final_total


@dataclass
class PrecheckRow:
    folder: str
    student_name: str
    class_name: str
    status: str
    issues: str
    typing_file: str
    resume_file: str
    can_grade: bool


# 自动生成的防覆盖后缀：_20261029_150737 或叠加重试 _20261029_150737_1
_AUTO_TS_SUFFIX = re.compile(r"(?:_\d{8}_\d{6})+(?:_\d+)?$")


def strip_auto_timestamp_stem(stem: str) -> str:
    """去掉文件名末尾自动加的时间戳，避免二次批改叠成 成绩表_时间_时间.xlsx。"""
    cleaned = _AUTO_TS_SUFFIX.sub("", stem or "").rstrip("_")
    return cleaned or "成绩表"


def unique_output_path(path: Path) -> Path:
    """
    生成不覆盖旧表的输出路径。
    - 若路径已带本程序自动时间戳（上次批改写入栏里的那种），先剥掉再换新时间戳
    - 若裸文件名且尚不存在，保持原名
    - 若裸文件名已存在，则改为 原名_年月日_时分秒.xlsx
    """
    path = Path(path)
    suffix = path.suffix if path.suffix else ".xlsx"
    raw_stem = path.stem
    clean_stem = strip_auto_timestamp_stem(raw_stem)
    had_auto_ts = clean_stem != raw_stem

    if not path.exists() and not had_auto_ts:
        return path if path.suffix else path.with_suffix(suffix)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = path.with_name(f"{clean_stem}_{stamp}{suffix}")
    n = 1
    while candidate.exists():
        candidate = path.with_name(f"{clean_stem}_{stamp}_{n}{suffix}")
        n += 1
    return candidate


def suggest_output_path(directory: Optional[Path] = None, class_hint: str = "") -> Path:
    """默认输出名：成绩表_班级_时间戳.xlsx（班级可空）。"""
    base = Path(directory) if directory else app_dir()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    hint = re.sub(r'[\\/:*?"<>|]+', "", (class_hint or "").strip())
    stem = f"成绩表_{hint}" if hint else "成绩表"
    # 直接拼时间戳；若同秒冲突由 unique 处理（会先剥旧戳再换新戳）
    return unique_output_path(base / f"{stem}_{stamp}.xlsx")


def extract_seat_or_id(folder_name: str) -> str:
    """从文件夹名尽量抠座位号/学号样片段，供教务简表预留；抠不到则空。"""
    raw = folder_name or ""
    if raw.startswith("(根目录)"):
        raw = raw[len("(根目录)") :]
    m = re.match(r"^([A-Za-z]\d{1,3})_", raw)
    if m:
        return m.group(1).upper()
    m = re.search(r"(?:学号|编号)[:：\s]*([A-Za-z0-9]{4,20})", raw)
    if m:
        return m.group(1)
    return ""


def unpack_submission_zip(zip_path: Path, dest_dir: Optional[Path] = None) -> Path:
    """
    把作业 ZIP 解到同级「原名_解压」目录（已存在则加时间戳）。
    仅支持 .zip；解压后返回学生作业总目录。
    """
    zip_path = Path(zip_path)
    if not zip_path.is_file() or zip_path.suffix.lower() != ".zip":
        raise ValueError("请选择 .zip 压缩包。")
    if dest_dir is None:
        dest_dir = zip_path.parent / f"{zip_path.stem}_解压"
        if dest_dir.exists():
            dest_dir = zip_path.parent / (
                f"{zip_path.stem}_解压_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as zf:
        # 尽量正确处理中文文件名（部分 Windows 打的包是 GBK）
        for info in zf.infolist():
            name = info.filename
            try:
                name_bytes = name.encode("cp437")
                try:
                    name = name_bytes.decode("gbk")
                except UnicodeDecodeError:
                    name = name_bytes.decode("utf-8", errors="replace")
            except Exception:
                pass
            target = dest_dir / name
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
    return dest_dir


def sort_by_name_pinyin(rows: list, name_attr: str = "student_name", folder_attr: str = "folder"):
    """按姓名拼音升序；同音再比姓名、文件夹。兼容 GradeRow / PrecheckRow。"""

    def key(row):
        name = getattr(row, name_attr, "") or ""
        folder = getattr(row, folder_attr, "") or ""
        py = "".join(lazy_pinyin(name)).lower()
        return (py, name, folder)

    return sorted(rows, key=key)


def build_precheck(students: list[StudentWork]) -> list[PrecheckRow]:
    rows: list[PrecheckRow] = []
    for sw in students:
        issues: list[str] = []
        has_t = sw.typing_path is not None
        has_r = sw.resume_path is not None
        if not has_t:
            issues.append("缺打字测试")
        if not has_r:
            issues.append("缺简历表")
        for note in sw.notes:
            if "快捷方式" in note:
                issues.append("含无效快捷方式(.lnk)")
            if "试卷原件" in note or "是试卷" in note:
                issues.append("交了试卷原件")
            if "转换失败" in note:
                issues.append(".doc转换失败")
            if "多份" in note:
                issues.append("同题多文件")
        # 去重并保持顺序
        seen = set()
        uniq = []
        for x in issues:
            if x not in seen:
                seen.add(x)
                uniq.append(x)
        if has_t and has_r and not uniq:
            status = "可批改"
        elif has_t or has_r:
            status = "部分可批（缺题）"
        else:
            status = "无法批改（两题皆缺）"
        can_grade = has_t or has_r
        rows.append(
            PrecheckRow(
                folder=sw.folder_name,
                student_name=sw.student_name or "",
                class_name=sw.class_name or "",
                status=status,
                issues="；".join(uniq) if uniq else "无",
                typing_file=sw.typing_path.name if sw.typing_path else "",
                resume_file=sw.resume_path.name if sw.resume_path else "",
                can_grade=can_grade,
            )
        )
    return sort_by_name_pinyin(rows, folder_attr="folder")


def precheck_submissions(submissions: Path) -> tuple[list[StudentWork], list[PrecheckRow]]:
    students = discover_students(Path(submissions))
    return students, build_precheck(students)


def deduction_stats(rows: list[GradeRow]) -> list[tuple[str, int, float]]:
    """返回 [(扣分点名称, 未满分人数, 占比%), ...] 按人数降序。"""
    if not rows:
        return []
    n = len(rows)
    keys = [
        ("打字完整度未满分", lambda r: r.typing_complete < SCORE_TYPING_COMPLETE - 0.05),
        ("打字准确率未满分", lambda r: r.typing_accuracy < SCORE_TYPING_ACCURACY - 0.05),
        ("简历完整度未满分", lambda r: r.resume_complete < 9.95),
        ("标题样式未满分", lambda r: r.resume_title < 4.95),
        ("字体/加粗未满分", lambda r: r.resume_font < 4.95),
        ("照片纵向/居中未满分", lambda r: r.resume_photo < 9.95),
        ("双线边框未满分", lambda r: r.resume_double < 9.95),
        ("灰色底纹未满分", lambda r: r.resume_gray < 9.95),
        ("文件命名未满分", lambda r: r.naming < 9.95),
        ("缺打字文件", lambda r: not r.typing_file),
        ("缺简历文件", lambda r: not r.resume_file),
    ]
    out: list[tuple[str, int, float]] = []
    for label, pred in keys:
        c = sum(1 for r in rows if pred(r))
        if c:
            out.append((label, c, round(100.0 * c / n, 1)))
    out.sort(key=lambda x: (-x[1], x[0]))
    return out


def _style_header_row(ws, headers: list[str], fill_color: str = "D9E2F3"):
    from openpyxl.utils import get_column_letter

    header_font = Font(bold=True)
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    fill = PatternFill("solid", fgColor=fill_color)
    for col, _ in enumerate(headers, 1):
        cell = ws.cell(1, col)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin
    return thin


def write_excel(
    rows: list[GradeRow],
    out: Path,
    precheck: Optional[list[PrecheckRow]] = None,
):
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "成绩表"
    headers = [
        "班级",
        "姓名",
        "文件夹",
        "打字文件",
        "简历文件",
        "打字完整度/20",
        "打字准确率/20",
        "打字小计/40",
        "简历完整度/10",
        "标题样式/5",
        "字体加粗/5",
        "照片纵向/10",
        "双线边框/10",
        "灰色底纹/10",
        "简历小计/50",
        "命名/10",
        "自动总分/100",
        "最终总分/100",
        "改分说明",
        "备注",
    ]
    ws.append(headers)
    thin = _style_header_row(ws, headers)

    for r in rows:
        final = r.effective_total()
        ws.append(
            [
                r.class_name,
                r.student_name,
                r.folder,
                r.typing_file,
                r.resume_file,
                r.typing_complete,
                r.typing_accuracy,
                r.typing_total,
                r.resume_complete,
                r.resume_title,
                r.resume_font,
                r.resume_photo,
                r.resume_double,
                r.resume_gray,
                r.resume_total,
                r.naming,
                r.total,
                final,
                r.adjust_note,
                r.remarks,
            ]
        )

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [
        16, 10, 28, 22, 22,
        12, 12, 12, 12, 10, 10, 10, 10, 10, 12, 10,
        12, 12, 18, 80,
    ]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"

    # —— 预检 / 问题名单 ——
    ws_pre = wb.create_sheet("预检报告")
    pre_headers = [
        "文件夹",
        "姓名",
        "班级",
        "状态",
        "问题",
        "打字文件",
        "简历文件",
        "可否批改",
    ]
    ws_pre.append(pre_headers)
    thin_pre = _style_header_row(ws_pre, pre_headers, "FCE4D6")
    pre_rows = precheck
    if pre_rows is None:
        # 从成绩行反推简版预检（成绩表已按拼音排，这里再统一排一次更稳）
        pre_rows = []
        for r in rows:
            issues = []
            if not r.typing_file:
                issues.append("缺打字测试")
            if not r.resume_file:
                issues.append("缺简历表")
            status = (
                "可批改"
                if r.typing_file and r.resume_file
                else ("部分可批（缺题）" if (r.typing_file or r.resume_file) else "无法批改（两题皆缺）")
            )
            pre_rows.append(
                PrecheckRow(
                    folder=r.folder,
                    student_name=r.student_name,
                    class_name=r.class_name,
                    status=status,
                    issues="；".join(issues) if issues else "无",
                    typing_file=r.typing_file,
                    resume_file=r.resume_file,
                    can_grade=bool(r.typing_file or r.resume_file),
                )
            )
    pre_rows = sort_by_name_pinyin(list(pre_rows), folder_attr="folder")
    warn_fill = PatternFill("solid", fgColor="F8CBAD")
    bad_fill = PatternFill("solid", fgColor="FF6B6B")
    for p in pre_rows:
        ws_pre.append(
            [
                p.folder,
                p.student_name,
                p.class_name,
                p.status,
                p.issues,
                p.typing_file,
                p.resume_file,
                "是" if p.can_grade else "否",
            ]
        )
        row_idx = ws_pre.max_row
        fill = None
        if not p.can_grade:
            fill = bad_fill
        elif p.status != "可批改":
            fill = warn_fill
        if fill:
            for col in range(1, len(pre_headers) + 1):
                ws_pre.cell(row_idx, col).fill = fill
    for row in ws_pre.iter_rows(min_row=2, max_row=ws_pre.max_row, max_col=len(pre_headers)):
        for cell in row:
            cell.border = thin_pre
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate([28, 10, 16, 18, 36, 22, 22, 10], 1):
        ws_pre.column_dimensions[get_column_letter(i)].width = w
    ws_pre.auto_filter.ref = ws_pre.dimensions
    ws_pre.freeze_panes = "A2"

    # —— 学情统计 ——
    ws_stat = wb.create_sheet("学情统计")
    totals = [r.effective_total() for r in rows]
    auto_totals = [r.total for r in rows]
    ws_stat["A1"] = "学情概览"
    ws_stat["A1"].font = Font(bold=True, size=14)
    ws_stat["A2"] = "人数"
    ws_stat["B2"] = len(rows)
    if totals:
        ws_stat["A3"] = "最终总分·平均"
        ws_stat["B3"] = round(sum(totals) / len(totals), 1)
        ws_stat["A4"] = "最终总分·最高"
        ws_stat["B4"] = max(totals)
        ws_stat["A5"] = "最终总分·最低"
        ws_stat["B5"] = min(totals)
        ws_stat["A6"] = "自动总分·平均"
        ws_stat["B6"] = round(sum(auto_totals) / len(auto_totals), 1)
        full = sum(1 for t in totals if t >= 99.5)
        pass_n = sum(1 for t in totals if t >= 60)
        fail_n = sum(1 for t in totals if t < 60)
        ws_stat["A7"] = "满分(≥99.5)"
        ws_stat["B7"] = full
        ws_stat["A8"] = "及格(≥60)"
        ws_stat["B8"] = pass_n
        ws_stat["A9"] = "不及格(<60)"
        ws_stat["B9"] = fail_n

    ws_stat["A11"] = "分数段分布（最终总分）"
    ws_stat["A11"].font = Font(bold=True)
    ws_stat["A12"] = "分数段"
    ws_stat["B12"] = "人数"
    ws_stat["C12"] = "占比%"
    bands = [
        ("90–100", 90, 100.01),
        ("80–89.9", 80, 90),
        ("70–79.9", 70, 80),
        ("60–69.9", 60, 70),
        ("0–59.9", 0, 60),
    ]
    for i, (label, lo, hi) in enumerate(bands):
        c = sum(1 for t in totals if lo <= t < hi)
        ws_stat.cell(13 + i, 1, label)
        ws_stat.cell(13 + i, 2, c)
        ws_stat.cell(13 + i, 3, round(100.0 * c / len(totals), 1) if totals else 0)

    ws_stat["A19"] = "高频未满分项（讲评用）"
    ws_stat["A19"].font = Font(bold=True)
    ws_stat["A20"] = "项目"
    ws_stat["B20"] = "未满分人数"
    ws_stat["C20"] = "占比%"
    for i, (label, c, pct) in enumerate(deduction_stats(rows)):
        ws_stat.cell(21 + i, 1, label)
        ws_stat.cell(21 + i, 2, c)
        ws_stat.cell(21 + i, 3, pct)
    ws_stat.column_dimensions["A"].width = 28
    ws_stat.column_dimensions["B"].width = 14
    ws_stat.column_dimensions["C"].width = 10

    # —— 教务简表（暂无固定模板：稳定列名，日后可映射）——
    ws_edu = wb.create_sheet("教务简表")
    edu_headers = ["班级", "姓名", "自动总分", "最终总分", "改分说明", "备注摘要"]
    ws_edu.append(edu_headers)
    thin_edu = _style_header_row(ws_edu, edu_headers, "E2EFDA")
    for r in rows:
        brief = (r.remarks or "").replace("\n", " / ")
        if len(brief) > 80:
            brief = brief[:77] + "…"
        ws_edu.append(
            [
                r.class_name,
                r.student_name,
                r.total,
                r.effective_total(),
                r.adjust_note,
                brief,
            ]
        )
    for row in ws_edu.iter_rows(min_row=2, max_row=ws_edu.max_row, max_col=len(edu_headers)):
        for cell in row:
            cell.border = thin_edu
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate([18, 10, 10, 10, 18, 60], 1):
        ws_edu.column_dimensions[get_column_letter(i)].width = w
    ws_edu.auto_filter.ref = ws_edu.dimensions
    ws_edu.freeze_panes = "A2"

    # —— 说明 ——
    ws2 = wb.create_sheet("说明")
    ws2["A1"] = "评分依据见 评分标准.md；本表由自动批改程序生成，异常项请人工抽查备注列。"
    ws2["A2"] = "标准答案：打字测试.docx / 简历表.docx"
    ws2["A3"] = f"作者：{AUTHOR}"
    ws2["A4"] = "「最终总分」默认同「自动总分」；教师复核改分后只改最终总分，并填写改分说明。"
    ws2["A5"] = "「教务简表」列名固定，便于以后对接教务模板（有固定表头后再映射）。"
    ws2["A6"] = "「预检报告」红色=两题皆缺无法批；橙色=缺题或部分问题，仍可能批出内容分。"
    ws2.column_dimensions["A"].width = 100

    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def most_common_class(students: list[StudentWork]) -> str:
    """从文件夹名解析出的班级里，取出现次数最多的；没有则空字符串。"""
    names = [sw.class_name.strip() for sw in students if sw.class_name.strip()]
    if not names:
        return ""
    return Counter(names).most_common(1)[0][0]


def grade_student(
    sw: StudentWork,
    answer_text: str,
    class_name: str,
) -> GradeRow:
    """批改单个学生，返回一行成绩（含路径，便于复核打开文件）。"""
    naming, naming_notes = score_naming(sw)
    tc, ta, tnotes = score_typing(sw.typing_path, answer_text)
    rs = score_resume(sw.resume_path)
    typing_total = round_half(tc + ta)
    total = round_half(typing_total + rs.total + naming)
    if naming >= 9.5 and not naming_notes:
        naming_notes = [
            "【文件命名】文件夹与文件名符合要求，本大题满分10分。"
        ]
    remark_parts = [x for x in (sw.notes + tnotes + rs.notes + naming_notes) if x]
    if not remark_parts:
        remark_parts = ["打字、简历、命名各项检查均达标，总分满分。"]
    elif total >= 99.5:
        remark_parts.append("【总评】三项合计满分100分。")
    remarks = "\n\n".join(remark_parts)
    return GradeRow(
        class_name=class_name,
        student_name=sw.student_name or "",
        folder=sw.folder_name,
        typing_file=sw.typing_path.name if sw.typing_path else "",
        resume_file=sw.resume_path.name if sw.resume_path else "",
        typing_complete=tc,
        typing_accuracy=ta,
        typing_total=typing_total,
        resume_complete=rs.completeness,
        resume_title=rs.title,
        resume_font=rs.font_bold,
        resume_photo=rs.photo,
        resume_double=rs.double_border,
        resume_gray=rs.gray,
        resume_total=rs.total,
        naming=naming,
        total=total,
        remarks=remarks,
        student_id=extract_seat_or_id(sw.folder_name),
        folder_path=str(sw.folder_path),
        typing_fullpath=str(sw.typing_path) if sw.typing_path else "",
        resume_fullpath=str(sw.resume_path) if sw.resume_path else "",
        final_total=total,
        adjust_note="",
    )


def load_typing_answer_text() -> str:
    typing_ans = resolve_answer("打字测试.docx")
    resume_ans = resolve_answer("简历表.docx")
    if not typing_ans.exists() or not resume_ans.exists():
        raise FileNotFoundError(
            f"缺少答案文件。请确认存在：\n  {typing_ans}\n  {resume_ans}"
        )
    return para_full_text(Document(str(typing_ans)))


def grade_all(
    submissions: Path,
    output: Path,
    class_name: Optional[str] = None,
    on_progress: Optional[Callable] = None,
    avoid_overwrite: bool = True,
) -> list[GradeRow]:
    """
    class_name:
      - None  → 自动取文件夹名里出现最多的班级；识别不到则留空
      - ""    → 强制留空
      - 其它  → 整表统一用该名称
    on_progress: 可选回调 (index_1based, total, display_name)
    avoid_overwrite: 若输出路径已存在，自动换带时间戳的新文件名
    """
    answer_text = load_typing_answer_text()

    students = discover_students(submissions)
    if not students:
        raise FileNotFoundError(
            f"在「{submissions}」下没有找到可批改的学生文件夹。\n"
            "请选择「一人一个子文件夹」的作业总目录（不要选程序自带的答案目录）。"
        )
    precheck = build_precheck(students)
    if class_name is None:
        class_name = most_common_class(students)
    rows: list[GradeRow] = []
    total_n = len(students)

    for idx, sw in enumerate(students, 1):
        if on_progress is not None:
            try:
                on_progress(idx, total_n, sw.student_name or sw.folder_name)
            except Exception:
                pass
        rows.append(grade_student(sw, answer_text, class_name or ""))

    rows = sort_by_name_pinyin(rows, folder_attr="folder")
    out = Path(output)
    if avoid_overwrite:
        out = unique_output_path(out)
    write_excel(rows, out, precheck=precheck)
    # 把实际路径写回，便于 GUI 打开
    grade_all.last_output = out  # type: ignore[attr-defined]
    return rows


def regrade_one_in_rows(
    rows: list[GradeRow],
    folder: str,
    submissions: Path,
    answer_text: Optional[str] = None,
    class_name: Optional[str] = None,
) -> list[GradeRow]:
    """按文件夹名重批一人，保留其他人的最终分/改分说明。"""
    if answer_text is None:
        answer_text = load_typing_answer_text()
    students = {sw.folder_name: sw for sw in discover_students(submissions)}
    sw = students.get(folder)
    if sw is None:
        raise FileNotFoundError(f"找不到学生文件夹：{folder}")
    # 班级：优先沿用原表该行；否则用传入/自动
    old = next((r for r in rows if r.folder == folder), None)
    cls = class_name
    if cls is None:
        cls = old.class_name if old else most_common_class(list(students.values()))
    new_row = grade_student(sw, answer_text, cls or "")
    if old and old.adjust_note:
        # 重批后默认恢复自动分，但保留历史改分说明作参考
        new_row.adjust_note = f"（已重批，原改分说明：{old.adjust_note}）"
    out_rows = []
    replaced = False
    for r in rows:
        if r.folder == folder:
            out_rows.append(new_row)
            replaced = True
        else:
            out_rows.append(r)
    if not replaced:
        out_rows.append(new_row)

    return sort_by_name_pinyin(out_rows, folder_attr="folder")


def save_rows(
    rows: list[GradeRow],
    output: Path,
    submissions: Optional[Path] = None,
    avoid_overwrite: bool = False,
) -> Path:
    """保存成绩（含复核改分后回写）。可选附带最新预检。"""
    precheck = None
    if submissions is not None and Path(submissions).is_dir():
        try:
            precheck = build_precheck(discover_students(Path(submissions)))
        except Exception:
            precheck = None
    out = Path(output)
    if avoid_overwrite:
        out = unique_output_path(out)
    write_excel(rows, out, precheck=precheck)
    return out


def self_check():
    """用标准答案自测，应接近满分（命名按答案文件夹不计）。"""
    typing_ans = resolve_answer("打字测试.docx")
    resume_ans = resolve_answer("简历表.docx")
    print("=== 自测：标准答案内容分 ===")
    print(f"打字答案: {typing_ans}")
    print(f"简历答案: {resume_ans}")
    answer_text = para_full_text(Document(str(typing_ans)))
    tc, ta, notes = score_typing(typing_ans, answer_text)
    print(f"打字: 完整度={tc}, 准确率={ta}, notes={notes}")
    assert tc == 20.0, tc
    assert ta == 20.0, ta

    rs = score_resume(resume_ans)
    print(
        f"简历: 完整={rs.completeness} 标题={rs.title} 字体={rs.font_bold} "
        f"照片={rs.photo} 双线={rs.double_border} 灰底={rs.gray} 合计={rs.total}"
    )
    print(f"notes={rs.notes}")
    assert rs.completeness >= 9.5, rs.completeness
    assert rs.title >= 4.5, rs.title
    assert rs.font_bold >= 4.0, rs.font_bold
    assert rs.photo >= 9.5, rs.photo
    assert rs.double_border >= 9.5, rs.double_border
    assert rs.gray >= 9.5, rs.gray
    assert rs.total >= 48.0, rs.total
    print("自测通过")


_NO_PAUSE = False


def _pause_if_frozen():
    """双击 exe 时避免窗口一闪而过。"""
    if _NO_PAUSE:
        return
    if getattr(sys, "frozen", False) and sys.stdin.isatty():
        try:
            input("\n按回车键退出…")
        except EOFError:
            pass


def _pick_submissions_dir() -> Path:
    """每次让用户选择学生提交文件夹，不设默认目录。"""
    base = app_dir()
    candidates = [
        p
        for p in sorted(base.iterdir())
        if p.is_dir()
        and p.name not in _SKIP_DIR_NAMES
        and not p.name.startswith(".")
    ]
    print(f"当前目录：{base}")
    print("请选择学生提交文件夹（一人一个子文件夹的那个总目录）：")
    if candidates:
        for i, p in enumerate(candidates, 1):
            print(f"  {i}. {p.name}")
        print("  0. 手动输入完整路径")
        raw = input("请输入序号：").strip()
        if not raw:
            raise SystemExit("未选择文件夹，已取消。")
        if raw == "0":
            typed = input("请输入学生提交文件夹的完整路径：").strip().strip('"')
            if not typed:
                raise SystemExit("未输入路径，已取消。")
            return Path(typed)
        try:
            idx = int(raw)
            if 1 <= idx <= len(candidates):
                return candidates[idx - 1]
        except ValueError:
            pass
        raise SystemExit(f"无效序号：{raw}")
    typed = input("旁边没有可选文件夹。请输入学生提交文件夹的完整路径：").strip().strip('"')
    if not typed:
        raise SystemExit("未输入路径，已取消。")
    return Path(typed)


def main():
    # Windows 控制台尽量用 UTF-8，减少中文乱码
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description=f"{APP_TITLE}（作者：{AUTHOR}）"
    )
    parser.add_argument(
        "--submissions",
        type=Path,
        default=None,
        help="学生提交根目录",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="成绩表输出路径",
    )
    parser.add_argument(
        "--class-name",
        default=None,
        help="成绩单「班级」列；不传则自动取文件夹名里出现最多的班级",
    )
    parser.add_argument(
        "--no-class",
        action="store_true",
        help="班级列留空（不自动识别）",
    )
    parser.add_argument("--self-check", action="store_true", help="仅跑答案自测")
    parser.add_argument(
        "--no-self-check",
        action="store_true",
        help="跳过开头的答案自测（加快批改）",
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="结束后不暂停（脚本调用时用）",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="打开图形界面",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="强制使用命令行（默认双击 exe 走图形界面）",
    )
    parser.add_argument(
        "--precheck",
        action="store_true",
        help="只做提交预检（缺题/快捷方式等），不写成绩、不改分",
    )
    parser.add_argument(
        "--allow-overwrite",
        action="store_true",
        help="允许覆盖已存在的成绩表（默认自动换带时间戳的新文件名）",
    )
    args = parser.parse_args()
    global _NO_PAUSE
    _NO_PAUSE = bool(args.no_pause)

    # 双击 exe 且未给批改参数时，默认打开 GUI
    want_gui = args.gui or (
        getattr(sys, "frozen", False)
        and not args.cli
        and not args.self_check
        and args.submissions is None
    )
    if want_gui:
        from grade_gui import run_gui

        run_gui()
        return

    try:
        if args.self_check:
            self_check()
            return

        print("=" * 50)
        print(f"  {APP_TITLE}")
        print(f"  作者：{AUTHOR}")
        print("=" * 50)
        print(f"程序目录：{app_dir()}")

        if args.submissions:
            submissions = args.submissions
        elif sys.stdin.isatty():
            submissions = _pick_submissions_dir()
        else:
            raise SystemExit(
                "未指定提交目录。请加参数：--submissions \"文件夹路径\""
            )
        if not submissions.exists():
            raise FileNotFoundError(f"学生提交目录不存在：{submissions}")

        if args.precheck:
            _students, pre = precheck_submissions(submissions)
            print(f"提交目录：{submissions}")
            print(f"预检人数：{len(pre)}")
            bad = [p for p in pre if not p.can_grade]
            warn = [p for p in pre if p.can_grade and p.status != "可批改"]
            print(f"可批改：{sum(1 for p in pre if p.status == '可批改')}")
            print(f"部分问题：{len(warn)}")
            print(f"无法批改：{len(bad)}")
            for p in pre:
                if p.status != "可批改":
                    print(f"  · {p.folder} | {p.status} | {p.issues}")
            print("预检完成（未写成绩表）。")
            return

        output = args.output or suggest_output_path(app_dir())
        if args.no_class:
            class_name: Optional[str] = ""
        elif args.class_name is not None:
            class_name = args.class_name.strip()
        else:
            class_name = None  # 批改时自动取出现最多的班级

        print(f"提交目录：{submissions}")
        print(f"输出文件：{output}")
        if class_name is None:
            print("班级名称：自动（取文件夹名中出现最多的）")
        else:
            print(f"班级名称：{class_name if class_name else '（空）'}")
        print("-" * 50)

        if not args.no_self_check:
            self_check()
            print("-" * 50)

        rows = grade_all(
            submissions,
            output,
            class_name=class_name,
            avoid_overwrite=not args.allow_overwrite,
        )
        real_out = getattr(grade_all, "last_output", output)
        used_class = rows[0].class_name if rows else ""
        print(
            f"班级栏实际填写：{used_class if used_class else '（空，未能从文件夹名识别）'}"
        )
        print(f"已批改 {len(rows)} 人，输出: {real_out}")
        totals = [r.effective_total() for r in rows]
        if totals:
            print(
                f"总分: 平均{sum(totals)/len(totals):.1f} "
                f"最高{max(totals)} 最低{min(totals)}"
            )
        stats = deduction_stats(rows)[:5]
        if stats:
            print("高频未满分：", "；".join(f"{a}×{b}" for a, b, _ in stats))
        print("完成。请打开 Excel 查看「成绩表 / 预检报告 / 学情统计 / 教务简表」。")
    except Exception as e:
        print(f"\n出错了：{e}")
        raise
    finally:
        _pause_if_frozen()


if __name__ == "__main__":
    main()
