# -*- coding: utf-8 -*-
"""
file-watcher-automation-template
================================
通用「文件到齐 → 调 WorkBuddy CLI 自动执行任务」监听器。

设计要点（每一条都是踩过坑之后的写法，改之前先读 README 的「六个坑」）：
  - 只读/移动本工程目录内的文件，不碰任何聊天软件进程。
  - 子进程必须 stdin=DEVNULL + CREATE_NO_WINDOW，否则 pythonw 下会弹黑窗并永久挂死。
  - CLI 必须强制精简产物（CODEBUDDY_FORCE_LITE_WB_BUNDLE=1），否则报 MODULE_NOT_FOUND。
  - CLI 的 bash 环境 PATH 与你的终端不同，python 一律用绝对路径。
  - 日志只回显参数，不回显提示词，避免自检关键字被自己污染。
  - 补日场景会并发触发多天任务，必须按日期升序串行，否则同写一个输出文件会互相覆盖。
  - --output-format text 重定向到文件是全缓冲，长任务中途看不到任何输出，改用 stream-json。
  - 判定「同一天的两份文件谁更新」不能只比大小：改一个数字字节数常常完全一致。
  - 「已产出就跳过」不等于一直对：源头改口径会让前面每一天都失效，见「报表周期与重算」。

目录约定：
  <project>/scripts/watcher.py     ← 本文件
  <project>/scripts/config.json    ← 配置（可复制 config.example.json 改名）
  <project>/<source dir>/...       ← 各源文件目录
  <project>/<log_dir>/             ← 运行日志（自动创建）
  <project>/<trigger_dir>/         ← ready 标记、锁（自动创建）
  <project>/<conclusion_dir>/      ← 每次任务完成的结论摘要（自动创建）
"""

import os
import re
import sys
import json
import time
import glob
import queue
import hashlib
import shutil
import threading
from datetime import datetime, timedelta
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent

try:
    from openpyxl import Workbook, load_workbook  # noqa: F401  (仅用于索引表，缺失不致命)
except Exception:
    Workbook = None
    load_workbook = None


# ---------------------------------------------------------------- 配置
DEFAULT_CONFIG = {
    "log_dir": "_logs",
    "trigger_dir": "_trigger",
    "conclusion_dir": "conclusions",
    "poll_seconds": 5,
    "stable_seconds": 3,
    # 补日窗口：往前追几天。源文件到位时间不固定时把它调大（如 7）。
    "lookback_days": 7,
    # 缺件等待日志的心跳间隔（秒）。整天没件时靠它证明监听还活着。
    "wait_log_interval_seconds": 1800,
    "auto_run": True,
    "preflight_on_start": True,
    "run_timeout": 5400,
    # ---- 确定性脚本优先（第七版，2026-10-24）----
    # 能用脚本算的就不该让模型算：脚本秒级、零 token、同输入同输出、可审计。
    # script_cmd 是命令模板列表，支持两个占位：{py} 解释器绝对路径、{day} 日期。
    #   例：["{py}", "scripts/gen.py", "daily", "{day}"]
    # 退出码契约（脚本自己实现）：0 成功｜2 口径不明转模型｜3 自检不过不落盘｜
    #                          4 当日已完成｜5 缺件。只有 2 会烧额度。
    "script_first": False,
    "script_timeout": 900,
    "script_cmd": [],
    # 本机 PATH 里没有 python 时的解释器绝对路径（留空则按内置候选顺序探测）
    "py_exe": "",
    # ---- 失败退避（第八版）----
    # 入队判断独立落盘，不再依赖内存里的 warned 集合：进程重启后仍然有效，
    # 也不会出现「某日入队失败过一次就永远不再入队、日志却写着下一轮重试」。
    "retry_backoff_seconds": [600, 1200, 1800, 3600],
    "retry_max": 12,
    # 产物判定：该日期已存在匹配文件时永不重跑，防止放宽窗口后推翻历史结果。
    # 支持 {y} {m} {d} 占位（不补零），例：output/报表{y}.{m}.{d}*.xlsx
    "produced_glob": "",
    "sources": [],
    # 外部来源（下载目录、聊天软件落盘目录等）：只复制不移动，按 mtime 取最新。
    "external_sources": [],
    "task": {"spec_doc": "", "prompt": ""},
    "cli": {
        "node": "",
        "entry": "",
        "git_bash": "",
        "product_version": "",
        # text（全缓冲，中途无输出）/ stream-json（实时流，推荐）
        "output_format": "stream-json",
        "perm_variants": [
            ["--permission-mode", "acceptEdits", "-y",
             "--allowedTools", "Bash(*),Read,Write,Edit,Glob,Grep"],
            ["-y"],
            ["--permission-mode", "bypassPermissions", "-y"],
        ],
    },
}

CONFIG_PATH = SCRIPT_DIR / "config.json"


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # 深拷贝
    if CONFIG_PATH.exists():
        try:
            user = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            for k, v in user.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as e:
            print(f"[!] 配置读取失败，改用默认值：{e}")
    return cfg


CFG = load_config()
LOGDIR = ROOT / CFG["log_dir"]
TRIGGER = ROOT / CFG["trigger_dir"]
CONC_DIR = ROOT / CFG.get("conclusion_dir", "conclusions")
STREAM_DIR = LOGDIR / "_stream"
HEARTBEAT = TRIGGER / "_heartbeat"   # 心跳文件，守护进程（watchdog.pyw）据此判断本进程是否还活着
for d in (LOGDIR, TRIGGER, CONC_DIR, STREAM_DIR):
    d.mkdir(parents=True, exist_ok=True)

IGNORE_PREFIX = ("~$", "._", ".")
# 判定「谁更新」时容许的 mtime 误差，避免复制动作本身带来的毫秒级抖动误判
MTIME_TOL = 2.0


# ---------------------------------------------------------------- 心跳
_last_beat = 0.0


def _beat(force: bool = False):
    """写心跳，供守护进程判断本进程是否还活着。

    为什么需要：无人值守场景下进程崩溃是无声的 —— 没有窗口、没有弹窗，
    日志停在最后一行，源文件照常往目录里落，但没人处理，要等下次重启
    才恢复。心跳 + 外部守护是唯一能覆盖这种情况的办法。

    节流到 60 秒一次：主循环轮询比这快得多，每轮写盘既没必要也不利于追溯。
    写盘失败一律吞掉 —— 监控自身绝不能影响主流程。
    """
    global _last_beat
    now = time.time()
    if not force and now - _last_beat < 60:
        return
    _last_beat = now
    try:
        HEARTBEAT.parent.mkdir(parents=True, exist_ok=True)
        HEARTBEAT.write_text(
            f"{os.getpid()}\n{datetime.now().isoformat(timespec='seconds')}\n",
            encoding="utf-8",
        )
    except Exception:
        pass


# ---------------------------------------------------------------- 工具
def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    try:
        with open(LOGDIR / f"watcher-{datetime.now():%Y-%m}.log", "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass


_STICKY_LAST: dict = {}
_STICKY_GAP = 300.0        # 同一标题 5 分钟内只弹一次，补跑多天时不刷屏


def _sticky_box(title: str, text: str, open_path=None):
    """不会自动消失的置顶结果框。

    为什么不用几秒就关的小窗：产物出没出、跑没跑成功，是「看一眼才知道
    能不能往下走」的信息。自动消失等于没提示，人不在屏幕前就错过了。
    这个框必须手动关，带「打开」按钮，省得再一层层点目录核对。

    判据只看标题里的字、不看调用方：带失败类字眼的用错误图标，其余信息图标。
    父窗口传 None —— pythonw 下没有控制台窗口。
    """
    now = time.time()
    if now - _STICKY_LAST.get(title, 0.0) < _STICKY_GAP:
        log(f"结果框节流跳过：{title}")
        return
    _STICKY_LAST[title] = now

    def _pop():
        try:
            import ctypes
            import subprocess
            bad = any(k in title for k in ("失败", "未产出", "异常", "不完整", "中止", "上限"))
            body = text
            flags = 0x00040000 | 0x00010000          # TOPMOST | SETFOREGROUND
            if open_path:
                body += "\n\n【是】打开　　【否】关闭"
                flags |= 0x00000004                   # YESNO
            else:
                body += "\n\n【确定】关闭"
            flags |= 0x00000010 if bad else 0x00000040  # ERROR / INFORMATION
            r = ctypes.windll.user32.MessageBoxW(None, body, title, flags)
            if open_path and r == 6:
                p = str(open_path)
                if os.path.isfile(p):
                    subprocess.Popen(["notepad.exe", p])
                else:
                    os.startfile(p)
        except Exception as e:
            log(f"结果框弹出失败：{title}｜{e}")

    threading.Thread(target=_pop, daemon=True).start()


def notify(title, text, sticky=False, open_path=None):
    """任务提示。按「错过之后能不能自己发现」分三档，别一律用最响的那档。

    sticky=True      置顶框，必须手动关 —— 产物出没出、跑成功没有，用这档
    notify_popup     几秒自动关的角标 —— 过程事件，漏看也无所谓
    notify_sound     只出声 + 写日志 —— 高频事件，否则会吵到关掉提示

    open_path 给了就在框里加「是/否」，点是直接打开文件或目录。
    """
    log(f"[通知] {title}｜{text}")
    if sticky:
        if CFG.get("notify_sticky", True):
            _sticky_box(title, text, open_path)
        return
    if CFG.get("notify_sound"):
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass


# ---------------------------------------------------------------- 弹窗自检（第九版，2026-10-24）
POPUP_PROBE = TRIGGER / "_popup_probe.txt"


def _check_popup_probe():
    """弹窗链路自检：往 <trigger>/_popup_probe.txt 写一行字即可。

    为什么需要它：想知道「弹窗到底弹不弹」，最自然的做法是随手起个进程弹一下，
    但会话/工具派生的进程常常活不过一分钟，窗口跟着进程一起消失，于是得出
    「弹不出来」的错误结论（2026-10-24 为此白查了两轮）。只有本进程弹的框，
    才代表无人值守时真会看到的那个框。

    用法：重启监听器后写入一行文字 → 下一轮轮询（默认 5 秒）按正式结果的
    弹法弹出该文字 → 弹完自动删除探针文件 → 日志留一条记录。
    """
    try:
        if not POPUP_PROBE.exists():
            return
        body = POPUP_PROBE.read_text(encoding="utf-8").strip() or "弹窗自检"
        POPUP_PROBE.unlink(missing_ok=True)
        log(f"弹窗自检：收到探针，弹出结果框｜{body[:60]}")
        notify("弹窗自检", body, sticky=True, open_path=ROOT)
    except Exception as e:
        log(f"弹窗自检失败：{e}")


def sha256(path, chunk=1 << 20):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            b = fh.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _is_stable(p: Path, min_age=None) -> bool:
    """文件必须落盘且静置数秒，避免读到「正在复制中」的半截文件"""
    min_age = CFG["stable_seconds"] if min_age is None else min_age
    try:
        st = p.stat()
    except OSError:
        return False
    if st.st_size <= 0:
        return False
    return (time.time() - st.st_mtime) >= min_age


def _resolve_dup(p: Path) -> Path:
    """兼容重复下载产生的 (1)（1）- 副本 等变体，找到后归一为标准名（不覆盖已有标准名）"""
    if p.exists() or not p.parent.exists():
        return p
    base = re.escape(p.stem)
    pat = re.compile(rf"^{base}\s*(?:[\(\（]\d+[\)\）]|-\s*副本(?:\s*[\(\（]\d+[\)\）]?)?)$")
    cands = [f for f in p.parent.iterdir()
             if f.is_file() and f.suffix.lower() == p.suffix.lower()
             and not f.name.startswith("~$") and pat.match(f.stem)]
    if not cands:
        return p
    src = max(cands, key=lambda f: f.stat().st_mtime)
    try:
        src.rename(p)
        log(f"文件名归一：{src.name} -> {p.name}")
    except Exception as e:
        log(f"文件名归一失败（{src.name}）：{e}")
        return src
    return p


# ---------------------------------------------------------------- 取件与产物判定
SKIP_SAME = "SKIP_SAME"     # 目标已在位且内容一致，无需动作
SKIP_STALE = "SKIP_STALE"   # 候选比在位版本更旧，弃用
OLD_DIR = LOGDIR / "_history"


def adopt(src: Path, dst: Path, copy_mode=True):
    """把候选文件放到目标位置，返回 (ok, note)。

    判定「谁是新版」的三层规则（只比大小的做法会把修正版当重复丢掉：
    改一个数字的修正版，字节数常常完全一致）：
      1. 大小不同 → 直接比 mtime
      2. 大小相同 → 比 SHA256；一致即真重复，返回 SKIP_SAME
      3. 内容不同 → 源 mtime 比在位版本新过 MTIME_TOL 才替换，否则返回 SKIP_STALE
    被替换的旧版不删除，移到 <log_dir>/_history/。
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    note = "新增"
    if dst.exists():
        same = False
        try:
            if dst.stat().st_size == src.stat().st_size:
                same = sha256(dst) == sha256(src)
        except OSError:
            same = False
        if same:
            return True, SKIP_SAME
        try:
            newer = src.stat().st_mtime > dst.stat().st_mtime + MTIME_TOL
        except OSError:
            newer = False
        if not newer:
            return True, SKIP_STALE
        OLD_DIR.mkdir(parents=True, exist_ok=True)
        old_ts = datetime.fromtimestamp(dst.stat().st_mtime).strftime("%Y%m%d-%H%M%S")
        bak = OLD_DIR / f"{dst.stem}.{old_ts}{dst.suffix}"
        shutil.move(str(dst), str(bak))
        note = f"替换旧版（旧版存 _history/{bak.name}）"
    if copy_mode:
        shutil.copy2(str(src), str(dst))   # copy2 保留 mtime，供下次比较
    else:
        shutil.move(str(src), str(dst))
    return True, note


def has_produced(day: str) -> bool:
    """该日期是否已经产出过结果文件（produced_glob 为空则视为不启用）。
    {y}{m}{d} 不补零，避免「10.1」与「10.11」撞号（后面若紧跟非数字字符即安全）。"""
    g = CFG.get("produced_glob") or ""
    if not g:
        return False
    try:
        y, m, d = [int(x) for x in day.split("-")]
        pat = g.format(y=y, m=m, d=d)
    except Exception:
        return False
    base = pat if os.path.isabs(pat) else str(ROOT / pat)
    return bool(glob.glob(base))


# ---------------------------------------------------------------- 源文件解析
DUP_SUFFIX = re.compile(r"\s*(?:[\(\（]\d+[\)\）]|-\s*副本(?:\s*[\(\（]\d+[\)\）]?)?)\s*$")


def strip_dup(name: str) -> str:
    """去掉重复下载产生的 (1)（1）- 副本，再做匹配。
    不做这一步的话，同名文件第二次下载成「X(1).xlsx」就匹配不上规则，
    「取 mtime 最新」也就永远看不到第二份了。"""
    stem, ext = os.path.splitext(name)
    return DUP_SUFFIX.sub("", stem) + ext


class Source:
    """一个源：目录 + 文件名正则。

    日期从哪来由 date_source 决定：
      "filename"（默认）：正则里用命名分组 y/m/d 取日期，y 可省略（按当年）
      "mtime"：正则只负责判断「是不是我这类型文件」，日期取文件的最后修改时间
               适合文件名固定、当天覆盖的报表（如「银行流水.xlsx」）
    """

    def __init__(self, spec, default_date_source="filename"):
        self.name = spec.get("name") or spec.get("dir", "source")
        self.dir = ROOT / spec["dir"]
        self.pattern = re.compile(spec["pattern"])
        self.allow_newer = bool(spec.get("allow_newer_fallback", False))
        self.std_name = spec.get("std_name", "")  # 可选：如 "A_{m}-{d}.xlsx"
        self.date_source = spec.get("date_source") or default_date_source
        # coverage：这份源「覆盖到哪天」。
        #   "cumulative" = 周期累计堆叠表（某天那份里含周期首日到该日的全部区块，
        #                  它一改，前面每一天都可能变）
        #   "single"（默认）= 只含当天数据
        # 修正版重算据此判断影响面，详见「报表周期与重算」一节。
        self.coverage = (spec.get("coverage") or "single").lower()

    def matches(self, name: str) -> bool:
        return bool(self.pattern.match(strip_dup(name)))

    def std_path(self, day: str) -> Path:
        y, m, d = [int(x) for x in day.split("-")]
        if self.std_name:
            return self.dir / self.std_name.format(y=y, m=m, d=d)
        return self.dir / f"{self.name}_{y}{m:02d}{d:02d}.xlsx"

    def parse_date(self, path: Path):
        """返回该文件归属的日期；不是本类型或不合法则返回 None"""
        name = strip_dup(path.name)
        if not self.pattern.match(name):
            return None
        if self.date_source == "mtime":
            try:
                return datetime.fromtimestamp(path.stat().st_mtime).date()
            except OSError:
                return None
        g = self.pattern.match(name).groupdict()
        try:
            now = datetime.now()
            y = int(g["y"]) if g.get("y") else now.year
            mo = int(g["m"])
            dd = int(g["d"])
            return datetime(y, mo, dd).date()
        except Exception:
            return None

    def exact(self, day: str):
        """当日文件。mtime 模式下按修改时间落在该日挑选，返回 None 表示没有"""
        if self.date_source != "mtime":
            p = _resolve_dup(self.std_path(day))
            return p if p.exists() else None
        day_d = datetime.fromisoformat(day).date()
        hits = [(f.stat().st_mtime, f) for d, f in self.list_all() if d == day_d]
        return max(hits, key=lambda t: t[0])[1] if hits else None

    def list_all(self):
        out = []
        if not self.dir.exists():
            return out
        for f in self.dir.iterdir():
            if not f.is_file() or f.name.startswith(IGNORE_PREFIX):
                continue
            dt = self.parse_date(f)
            if dt:
                out.append((dt, f))
        return sorted(out, key=lambda t: (t[0], t[1].stat().st_mtime))

    def pick(self, day: str):
        """返回 (path 或 None, 是否回溯, 实际日期)。
        当日缺失且 allow_newer_fallback 时，取日期更晚的最近一份（适合周期累计堆叠表）。"""
        day_d = datetime.fromisoformat(day).date()
        p = self.exact(day)
        if p is not None:
            return p, False, day_d
        if self.allow_newer:
            newer = [(d, f) for d, f in self.list_all() if d >= day_d]
            if newer:
                d, f = newer[0]
                return f, True, d
        return None, False, day_d


SOURCES = [Source(s, CFG.get("date_source", "filename"))
           for s in CFG.get("sources", [])]


# ---------------------------------------------------------------- 外部来源取件
# 场景：文件先落在别处（浏览器下载目录、聊天软件的文件目录），
# 使用者不会手动另存。此时只靠监听源目录永远等不到文件。
# 规则：按 Source 的文件名正则去外部目录里认领，复制到源目录，绝不移动/删除原件。
def _ext_cfgs():
    out = []
    for e in (CFG.get("external_sources") or []):
        if not e.get("enabled"):
            continue
        root = Path(e.get("root") or "")
        if not str(e.get("root") or "").strip():
            continue
        out.append({
            "name": e.get("name") or root.name or "external",
            "root": root,
            # depth = 从 root 往下最多走几层子目录（root 的直接子项算第 1 层）。
            # 微信 4.x 的文件在 <root>\<wxid>\msg\file\<YYYY-MM>\，需要 3 层以上。
            "depth": int(e.get("depth", 4)),
            "max_age_days": float(e.get("max_age_days", 8)),
            "interval_seconds": float(e.get("interval_seconds", 30)),
            "only": list(e.get("sources") or []),
        })
    return out


def scan_external(cfg):
    """返回 [(Source, path)]，只做匹配不做稳定性判断"""
    cutoff = time.time() - cfg["max_age_days"] * 86400
    targets = [s for s in SOURCES if not cfg["only"] or s.name in cfg["only"]]
    targets = [s for s in targets if s.date_source != "mtime"]  # mtime 源无法按日期落位
    out = []
    if not cfg["root"].exists() or not targets:
        return out

    def walk(d: Path, lvl: int):
        try:
            entries = list(d.iterdir())
        except OSError:
            return
        for p in entries:
            if p.is_dir():
                if lvl < cfg["depth"]:
                    walk(p, lvl + 1)
                continue
            if not p.is_file() or p.name.startswith(IGNORE_PREFIX):
                continue
            try:
                if p.stat().st_mtime < cutoff:
                    continue
            except OSError:
                continue
            for s in targets:
                if s.matches(p.name):
                    out.append((s, p))
                    break

    walk(cfg["root"], 1)
    return out


def dedupe_external(cands):
    """同一「源+日期」可能有多份（重复下载、对方发的修正版），只留 mtime 最新一份"""
    best = {}
    for s, p in cands:
        day = s.parse_date(p)
        if day is None:
            continue
        try:
            mt = p.stat().st_mtime
        except OSError:
            continue
        key = (s.name, day.isoformat())
        if key not in best or mt > best[key][0]:
            best[key] = (mt, s, p)
    return [(v[1], v[2]) for v in best.values()]


def pickup_external(cands):
    """复制外部候选文件到源目录。返回实际搬运数量"""
    n = 0
    for s, p in cands:
        day = s.parse_date(p)
        if day is None:
            continue
        if not _is_stable(p, min_age=CFG["stable_seconds"]):
            continue
        dst = s.std_path(day.isoformat())
        try:
            ok, note = adopt(p, dst, copy_mode=True)
        except Exception as e:
            log(f"外部取件失败 [{s.name}] {p.name}：{e}")
            continue
        if note in (SKIP_SAME, SKIP_STALE):
            continue
        try:
            rel = dst.relative_to(ROOT)
        except Exception:
            rel = dst
        log(f"外部取件 [{s.name}] {p.name} -> {rel}｜{note}")
        n += 1
    return n


def check_ready(day: str):
    """所有源都取到即齐备。返回 (ok, missing 列表, ready 标记路径)"""
    rp = TRIGGER / f"ready_{day}.json"
    picks, missing, unstable = [], [], []
    for s in SOURCES:
        p, fb, cov = s.pick(day)
        if p is None:
            missing.append(f"{s.name}（{s.dir.name}/ 当日无匹配文件）")
            continue
        if not _is_stable(p):
            unstable.append(f"{s.name}:{p.name}")
        picks.append({"source": s.name, "path": str(p), "fallback": fb,
                      "covers_to": cov.isoformat(),
                      "size": p.stat().st_size,
                      "mtime": datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec="seconds")})
    if missing or unstable:
        return False, missing + [f"{x}（正在写入）" for x in unstable], rp

    payload = {"date": day, "ready_at": datetime.now().isoformat(timespec="seconds"),
               "sources": picks}
    prev = None
    if rp.exists():
        try:
            prev = json.loads(rp.read_text(encoding="utf-8"))
        except Exception:
            prev = None
    if not (prev and prev.get("sources") == picks):
        rp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return True, [], rp


# ---------------------------------------------------------------- 失败退避（第八版，2026-10-24）
RETRY_DIR = TRIGGER / "_retry"


def _retry_file(day: str):
    return RETRY_DIR / f"{day}.json"


def _retry_gate(day: str):
    """现在能不能再入队。返回 (是否跑, 第几次尝试, 上次错误)。

    为什么落盘而不是放内存：2026-10-24 踩过的坑 —— 入队判断被 `warned` 包着，
    而 warned 在 while 外持久，某日只要入队过一次就再也不会重入。执行失败
    （典型是鉴权过期）就永久卡死，日志却一直写着「下一轮重试」。改成磁盘文件：
    跨进程重启有效，重启后不会把失败次数清零。
    """
    p = _retry_file(day)
    if not p.exists():
        return True, 1, ""
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return True, 1, ""
    n = int(d.get("attempts") or 0)
    if time.time() < float(d.get("next_ts") or 0):
        return False, n, str(d.get("last_error") or "")
    return True, n + 1, str(d.get("last_error") or "")


def _mark_fail(day: str, reason: str):
    """记一次失败并安排下次重试时间，返回 (第几次失败, 等待秒数)。"""
    steps = CFG.get("retry_backoff_seconds") or [600, 1200, 1800, 3600]
    cap = int(CFG.get("retry_max", 12))
    p = _retry_file(day)
    n = 0
    try:
        if p.exists():
            n = int(json.loads(p.read_text(encoding="utf-8")).get("attempts") or 0)
    except Exception:
        n = 0
    n += 1
    wait = steps[min(n - 1, len(steps) - 1)]
    if n > cap:
        wait = 6 * 3600          # 连续失败太多次：拉到 6 小时一探，避免整点刷屏
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps({
            "day": day,
            "attempts": n,
            "next_ts": time.time() + wait,
            "next_at": (datetime.now() + timedelta(seconds=wait)).isoformat(timespec="seconds"),
            "last_error": str(reason)[:200],
        }, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        log(f"失败记录写入异常（不影响主流程）：{e}")
    return n, wait


def _clear_fail(day: str):
    try:
        _retry_file(day).unlink(missing_ok=True)
    except Exception:
        pass


# ---------------------------------------------------------------- 脚本通道（第一路）
# 第七版：能确定性算出来的结果，不该让模型即兴算。脚本秒级、零 token、
# 同输入同输出、出错能复现；模型只负责「脚本明确说判断不了」的那部分。
_PY_CANDIDATES = [
    r"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.14.3\python.exe",
    r"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe",
]


def pick_python():
    """本机 PATH 常常没有 python 命令，必须用解释器绝对路径（配置 py_exe 可覆盖）"""
    p = CFG.get("py_exe")
    if p and Path(p).exists():
        return Path(p)
    for c in _PY_CANDIDATES:
        if Path(c).exists():
            return Path(c)
    return None


def _script_argv(day: str):
    """按 script_cmd 模板生成命令；{py} 换成解释器绝对路径，{day} 换成日期。"""
    raw = CFG.get("script_cmd") or []
    if not raw:
        return None
    py = pick_python()
    out = []
    for x in raw:
        s = str(x).replace("{day}", day)
        if "{py}" in s:
            s = s.replace("{py}", str(py) if py else "")
        out.append(s)
    return out


def run_script(day: str, out_log: Path, timeout: float):
    """跑确定性脚本，返回 (退出码, 输出全文, 启动失败原因)。

    退出码为 None 表示脚本根本没跑起来（没配 script_cmd、解释器或脚本缺失、
    超时、启动异常）。这类属环境问题，**不能**转交模型兜底 —— 兜底也跑不动，
    只会白烧额度，所以调用方要单独处理。
    """
    import subprocess

    argv = _script_argv(day)
    if not argv:
        return None, "", "未配置 script_cmd"
    if not argv[0]:
        return None, "", "找不到 python 解释器（script_cmd 用了 {py} 但本机未探测到）"

    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    t0 = time.time()
    try:
        proc = subprocess.run(
            [str(x) for x in argv],
            cwd=str(ROOT), env=env, stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            timeout=timeout,
        )
        rc = proc.returncode
        text = proc.stdout.decode("utf-8", "replace")
    except subprocess.TimeoutExpired as e:
        rc = None
        text = (e.stdout or b"").decode("utf-8", "replace")
        text += f"\n[超时] 超过 {timeout / 60:.0f} 分钟被终止\n"
    except Exception as e:
        return None, "", f"脚本启动失败：{e}"

    try:
        out_log.write_text(
            f"# {' '.join(str(x) for x in argv)}\n"
            f"# 退出码 {rc}｜耗时 {time.time() - t0:.1f} 秒｜"
            f"{datetime.now():%Y-%m-%d %H:%M:%S}\n\n{text}",
            encoding="utf-8",
        )
    except Exception:
        pass
    return rc, text, ""


def _last_line(text: str) -> str:
    for ln in reversed([x.strip() for x in (text or "").splitlines()]):
        if ln:
            return ln[:180]
    return "（无输出）"


def save_script_conclusion(day: str, rc, text, elapsed: float, event: str = "成功"):
    """脚本通道的结论落到 <conclusion_dir>/<day>.md，与模型通道同一位置。

    2026-10-24 修：原来只在成功时写，缺件/自检不过/脚本起不来这些情况在结论
    目录里整片空白，事后完全看不出那天发生了什么。现在每个终态都写一条
    （同一天覆盖写，不会堆积），「结论」才真的是运行日志而不是成功日志。
    """
    p = CONC_DIR / f"{day}.md"
    body = (text or "").strip() or "（脚本无输出）"
    rc_txt = "未启动" if rc is None else str(rc)
    try:
        CONC_DIR.mkdir(parents=True, exist_ok=True)
        p.write_text(
            f"# {day} 运行结论\n\n"
            f"- 完成时间：{datetime.now():%Y-%m-%d %H:%M:%S}\n"
            f"- 结果：{event}\n"
            f"- 通道：确定性脚本（未经模型）\n"
            f"- 耗时：{elapsed:.1f} 秒\n"
            f"- 退出码：{rc_txt}"
            f"（0 成功｜2 口径不明｜3 自检不过｜4 当日已完成｜5 缺件）\n\n"
            f"## 脚本输出\n\n```\n{body}\n```\n",
            encoding="utf-8")
    except Exception as e:
        log(f"结论文件写入异常（不影响主流程）：{e}")
    return p


# ---------------------------------------------------------------- CLI 调用
def build_prompt(day: str):
    tpl = (CFG.get("task") or {}).get("prompt") or ""
    spec = (CFG.get("task") or {}).get("spec_doc") or ""
    notes = []
    for s in SOURCES:
        p, fb, cov = s.pick(day)
        if fb and p is not None:
            notes.append(f"{s.name} 当日文件缺失，改用 {p.name}（覆盖至 {cov}）回溯取 {day} 这一段；"
                         f"禁止用其他日期顶替或外推，摘要中须注明「{s.name} 源为 {p.name} 回溯取数」。")
    return tpl.format(root=ROOT, spec=(ROOT / spec) if spec else "",
                      day=day, notes="；".join(notes) or "无")


def _cli_cfg():
    c = CFG.get("cli") or {}
    return {
        "node": c.get("node") or "node",
        "entry": c.get("entry") or "",
        "git_bash": c.get("git_bash") or "",
        "product_version": c.get("product_version") or "",
        "variants": c.get("perm_variants") or DEFAULT_CONFIG["cli"]["perm_variants"],
    }


def _cli_env():
    env = os.environ.copy()
    env["CODEBUDDY_FORCE_LITE_WB_BUNDLE"] = "1"
    env["CODEBUDDY_FORCE_HEADLESS_BUNDLE"] = "1"
    gb = _cli_cfg()["git_bash"]
    if gb:
        env.setdefault("CODEBUDDY_CODE_GIT_BASH_PATH", gb)
    env["CODEBUDDY_SKIP_GIT_BASH_CHECK"] = "1"
    pv = _cli_cfg()["product_version"]
    if pv:
        env.setdefault("CLIENT_INFO_PRODUCT_VERSION", pv)
    return env


def _stream_line_to_text(line: str) -> str:
    """把 stream-json 的一行事件转成人读文本。解析不了一律原样保留（stderr、
    错误堆栈、Node 警告都在这些行里，瞬时故障重试就是靠它们判定的）。"""
    s = line.rstrip("\r\n")
    if not s.strip():
        return ""
    try:
        ev = json.loads(s)
    except Exception:
        return s
    if not isinstance(ev, dict):
        return s
    t = ev.get("type")
    if t == "system":
        model = ev.get("model") or (ev.get("message") or {}).get("model") or "?"
        return f"[会话] {ev.get('subtype', 'init')}｜model={model}"
    if t == "assistant":
        out = []
        for c in ((ev.get("message") or {}).get("content") or []):
            if not isinstance(c, dict):
                continue
            if c.get("type") == "text" and c.get("text", "").strip():
                out.append(c["text"].rstrip())
            elif c.get("type") == "tool_use":
                inp = c.get("input") or {}
                brief = (inp.get("command") or inp.get("file_path")
                         or inp.get("pattern") or inp.get("description") or "")
                out.append(f"[工具] {c.get('name', '?')}｜{str(brief)[:130]}")
        return "\n".join(out)
    if t == "user":
        # tool_result 正常内容不记，只保留报错，避免日志被回显撑爆
        for c in ((ev.get("message") or {}).get("content") or []):
            if isinstance(c, dict) and c.get("type") == "tool_result" and c.get("is_error"):
                txt = c.get("content")
                if isinstance(txt, list):
                    txt = " ".join(x.get("text", "") for x in txt if isinstance(x, dict))
                return "[工具报错] " + str(txt)[:200]
        return ""
    if t == "result":
        parts = [f"[完成] subtype={ev.get('subtype', '?')}"]
        dur = ev.get("duration_ms")
        if isinstance(dur, (int, float)):
            parts.append(f"耗时 {dur / 1000:.1f}s")
        if ev.get("is_error"):
            parts.append("is_error=true")
        return "  ".join(parts)
    return f"[事件] type={t}"


def _run_cli(prompt, out_log: Path, timeout, title, extra_args=None):
    """统一 CLI 调用：无窗口 + stdin 断开 + 强制精简产物 + 实时流输出。

    日志只回显权限参数，绝不回显提示词（否则关键字会污染自检判定）。
    --output-format text 会被全缓冲，长任务中途看不到任何东西；
    改用 stream-json 后逐行读取、逐行 flush，跑动时就能看到工具调用。
    """
    import subprocess
    c = _cli_cfg()
    perm = list(extra_args or c["variants"][0])
    fmt = (CFG.get("cli") or {}).get("output_format") or "stream-json"
    cmd = [c["node"], c["entry"], "-p", prompt] + perm + ["--output-format", fmt]
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    STREAM_DIR.mkdir(parents=True, exist_ok=True)
    stream_log = STREAM_DIR / f"{Path(out_log).stem}-{datetime.now():%Y%m%d-%H%M%S}.jsonl"
    holder = {"result": ""}

    with open(out_log, "a", encoding="utf-8") as fh, \
            open(stream_log, "w", encoding="utf-8") as sf:
        fh.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} {title} =====\n")
        fh.write(f"[args] {' '.join(perm)}｜format={fmt}\n")
        fh.flush()
        proc = subprocess.Popen(
            cmd, cwd=str(ROOT), stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            env=_cli_env(), creationflags=flags,
            text=True, encoding="utf-8", errors="replace", bufsize=1)

        def pump():
            try:
                for line in proc.stdout:
                    sf.write(line)
                    sf.flush()
                    if fmt == "stream-json":
                        txt = _stream_line_to_text(line)
                        t = _assistant_text(line)      # 只取正文，供结论摘要用
                        if t:
                            holder["result"] = t
                    else:
                        txt = line.rstrip("\r\n")
                    if txt:
                        fh.write(txt + "\n")
                        fh.flush()
            except Exception:
                pass

        th = threading.Thread(target=pump, daemon=True)
        th.start()
        rc = None
        try:
            rc = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            rc = -9
        th.join(5)
        fh.write(f"[退出] returncode={rc}｜{datetime.now():%Y-%m-%d %H:%M:%S}\n")
        fh.flush()

    r = subprocess.CompletedProcess(cmd, rc)
    r.result_text = holder["result"]
    r.stream_log = stream_log
    return r


def _assistant_text(line: str) -> str:
    """取 assistant 事件的正文（不含工具调用），用于结论摘要"""
    try:
        ev = json.loads(line)
    except Exception:
        return ""
    if not isinstance(ev, dict) or ev.get("type") != "assistant":
        return ""
    parts = [c.get("text", "") for c in ((ev.get("message") or {}).get("content") or [])
             if isinstance(c, dict) and c.get("type") == "text"]
    return "\n".join(p.strip() for p in parts if p and p.strip())


def save_conclusion(day: str, r, elapsed: float, out_log: Path):
    """把任务结果写成 <conclusion_dir>/<日期>.md，人不用翻冗长日志"""
    try:
        body = (getattr(r, "result_text", "") or "").strip()
        lines = [
            f"# {day} 任务执行结论",
            "",
            f"- 完成时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
            f"- 耗时：{elapsed / 60:.1f} 分钟",
            f"- 退出码：{r.returncode}",
            f"- 输出日志：{out_log.name}",
            f"- 原始事件流：{getattr(r, 'stream_log', '') or '（无）'}",
            "",
            "## 结论",
            "",
            body or "（本次未取到 result 正文，请查看输出日志）",
            "",
        ]
        (CONC_DIR / f"{day}.md").write_text("\n".join(lines), encoding="utf-8")
    except Exception as e:
        log(f"结论写入失败：{e}")


TRANSIENT_HINTS = ("temporarily unavailable", "500", "502", "503", "504", "429",
                   "rate limit", "overloaded", "timeout", "ECONNRESET", "ETIMEDOUT")


def _looks_transient(out_log: Path, offset=0) -> bool:
    try:
        with open(out_log, "r", encoding="utf-8", errors="ignore") as fh:
            fh.seek(max(0, offset))
            chunk = fh.read()[-4000:].lower()
    except Exception:
        return False
    return any(h in chunk for h in TRANSIENT_HINTS)


def _run_cli_retry(prompt, out_log, timeout, title, extra_args=None,
                   attempts=3, waits=(45, 150)):
    """只对上游瞬时故障重试，业务失败不重试"""
    r = None
    for i in range(attempts):
        tag = title if i == 0 else f"{title}（第 {i + 1} 次尝试）"
        try:
            pos = out_log.stat().st_size
        except Exception:
            pos = 0
        r = _run_cli(prompt, out_log, timeout, tag, extra_args)
        if r.returncode == 0:
            return r
        if i < attempts - 1 and _looks_transient(out_log, pos):
            w = waits[min(i, len(waits) - 1)]
            log(f"{title} 遇上游瞬时故障（exit {r.returncode}），{w}s 后重试")
            time.sleep(w)
            continue
        return r
    return r


# ---------------------------------------------------------------- 只读预检
def run_preflight():
    """启动时自检：A 通道 → B 权限/python → C 由使用方自定义提示词验证。
    结果写 <log_dir>/preflight-{日期}.log。只读，不改业务文件。"""
    day = datetime.now().date().isoformat()
    out_log = LOGDIR / f"preflight-{day}.log"
    c = _cli_cfg()
    if not c["entry"] or not Path(c["entry"]).exists():
        log("预检跳过：未配置 cli.entry 或路径不存在")
        return
    try:
        out_log.write_text(f"===== {datetime.now():%Y-%m-%d %H:%M:%S} 只读预检 =====\n",
                           encoding="utf-8")
        log(f"预检 A/2：通道验证｜输出 {out_log.name}")
        r1 = _run_cli("Reply with exactly this token: PF_OK", out_log, 300, "阶段A 通道验证")
        if r1.returncode != 0:
            log(f"预检 A 失败（exit {r1.returncode}）｜详见 {out_log.name}")
            return
        log("预检 A 通过")

        probe = ("Use the Bash tool to run exactly this command, then copy its real stdout:\n"
                 "python -c \"import platform;print('PY_OK',platform.python_version())\"\n"
                 "Rules: if Bash is blocked/denied reply exactly SHELL_DENIED; "
                 "if it ran reply exactly one line 'PY_OK <version>' copied from stdout; "
                 "never invent versions.")
        chosen = None
        for perm in c["variants"]:
            name = " ".join(perm)[:48]
            log(f"预检 B/2：权限探测 [{name}]")
            try:
                pos = out_log.stat().st_size
            except Exception:
                pos = 0
            r = _run_cli_retry(probe, out_log, 600, f"阶段B 权限探测 {name}",
                               extra_args=perm, attempts=2, waits=(30,))
            try:
                with open(out_log, "r", encoding="utf-8", errors="ignore") as fh:
                    fh.seek(pos)
                    seg = fh.read()
            except Exception:
                seg = ""
            if r.returncode == 0 and re.search(r"(?m)^\s*PY_OK\s+\d", seg):
                chosen = perm
                (TRIGGER / "_perm_mode.json").write_text(json.dumps(
                    {"args": perm, "at": datetime.now().isoformat(timespec="seconds")},
                    ensure_ascii=False, indent=2), encoding="utf-8")
                log(f"预检 B 通过：权限组合 [{name}] 可执行命令，已记录 _perm_mode.json")
                break
            log(f"权限组合 [{name}] 未通过（exit {r.returncode}）")
        if chosen is None:
            log("预检失败：所有权限组合都无法执行命令｜详见日志")
            return
        log(f"预检完成｜详见 {out_log.name}")
    except Exception as e:
        log(f"预检异常：{e}")


def _perm_args():
    try:
        d = json.loads((TRIGGER / "_perm_mode.json").read_text(encoding="utf-8"))
        if isinstance(d.get("args"), list) and d["args"]:
            return d["args"]
    except Exception:
        pass
    return _cli_cfg()["variants"][0]


# ---------------------------------------------------------------- 执行任务
_RUN_SERIAL = threading.Lock()   # 手动单日调用时的互斥（主循环走队列）

# 待执行队列：补日窗口内可能同时有多天齐备，必须严格按日期升序逐个跑。
# 后一天产物依赖前一天产物，顺序错会算错累计数；靠「抢锁」无法保证顺序。
_RUN_QUEUE = queue.Queue()
_QUEUED = set()
# 缺件等待日志状态：day -> (缺件签名, 上次记录时间戳)
_MISS = {}


# ---------------------------------------------------------------- 临时脚本收纳
# CLI 跑任务时常常在 scripts 目录里现场写临时脚本（探测、生成、校验各一套），
# 一天一套，几个月就把正经文件埋了。所以每次任务结束把「本次新增」的文件
# 搬进 <log_dir>/_tmp/<日期>/ —— 只搬不删，复查还能翻到。
_DEFAULT_SWEEP_KEEP = [
    "watcher.py", "watchdog.pyw", "config.json", "config.example.json",
    "start_watcher.vbs", "startup_watcher.vbs",
]


def _sweep_cfg():
    c = CFG.get("scripts_sweep") or {}
    keep = set(_DEFAULT_SWEEP_KEEP) | set(c.get("keep") or [])
    return {
        "enabled": bool(c.get("enabled", True)),
        "dir": c.get("dir") or "_logs/_tmp",
        "keep": keep,
        "purge": bool(c.get("purge_enabled", True)),
        "keep_days": int(c.get("keep_days", 30)),
        "purge_hours": float(c.get("purge_interval_hours", 6)),
    }


def _rel_safe(p):
    try:
        return str(Path(p).relative_to(ROOT))
    except Exception:
        return str(p)


def snapshot_scripts():
    """任务开始前给 scripts 目录拍快照（只记文件名）；关闭时返回 None"""
    if not _sweep_cfg()["enabled"]:
        return None
    try:
        return {p.name for p in SCRIPT_DIR.iterdir() if p.is_file()}
    except Exception as ex:
        log(f"脚本快照失败（不影响任务）：{ex}")
        return None


def sweep_scripts(day: str, before):
    """把本次任务新产生的文件收进 _tmp/<日期>/：只搬不删，白名单一律不动

    只在任务**全部结束**后调用 —— 一次任务里可能先生成探测脚本、再生成、再校验，
    中途搬走会让它自己找不到文件。
    """
    if before is None:
        return 0
    sc = _sweep_cfg()
    if not sc["enabled"]:
        return 0
    try:
        now = {p.name for p in SCRIPT_DIR.iterdir() if p.is_file()}
    except Exception:
        return 0
    fresh = sorted(now - before - sc["keep"])
    if not fresh:
        return 0
    dest = ROOT / sc["dir"] / day
    moved = 0
    for name in fresh:
        try:
            dest.mkdir(parents=True, exist_ok=True)
            shutil.move(str(SCRIPT_DIR / name), str(dest / name))
            moved += 1
        except Exception as ex:
            log(f"临时文件收纳失败 {name}｜{ex}")
    if moved:
        log(f"临时文件已收纳：{moved} 个 -> {_rel_safe(dest)}")
    return moved


# _tmp 下的日期目录名必须严格长这样，其它一律不碰（防止误删人工放进去的东西）
_TMP_NAME_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def purge_tmp(force: bool = False) -> int:
    """清理 _tmp 下超过 keep_days 的日期目录，让收纳不会无限堆积。

    自动删除所以刻意保守：
      - 只处理 _tmp 的**直接子目录**，名字必须严格是 YYYY-MM-DD，认不出一律跳过；
      - 解析后再次校验路径正好在 _tmp 下一层，防配置写错指到别处；
      - 只删早于 keep_days 天的，今天的和所有非日期目录不动；
      - 全程 try/except，任何一步出错只记日志，不影响正在跑的任务；
      - 用 .last_purge 时间戳节流，默认 6 小时最多清一次。
    """
    try:
        sc = _sweep_cfg()
        if not sc["purge"]:
            return 0
        base = ROOT / sc["dir"]
        if not base.is_dir():
            return 0
        base_r = base.resolve()
        stamp = base / ".last_purge"
        if not force:
            try:
                if stamp.exists() and (time.time() - stamp.stat().st_mtime) < sc["purge_hours"] * 3600:
                    return 0
            except Exception:
                pass
        cutoff = datetime.now().date() - timedelta(days=max(1, sc["keep_days"]))
        removed = 0
        for d in list(base.iterdir()):
            try:
                if not d.is_dir() or not _TMP_NAME_RE.match(d.name):
                    continue
                try:
                    dd = datetime.strptime(d.name, "%Y-%m-%d").date()
                except ValueError:
                    continue
                if dd >= cutoff:
                    continue
                d_r = d.resolve()
                if d_r.parent != base_r:
                    log(f"临时文件清理跳过（路径异常）：{d}")
                    continue
                shutil.rmtree(d_r)
                removed += 1
                log(f"临时文件清理：移除 {d.name}（超过 {sc['keep_days']} 天）")
            except Exception as ex:
                log(f"临时文件清理跳过 {getattr(d, 'name', d)}｜{ex}")
        try:
            stamp.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
        except Exception:
            pass
        return removed
    except Exception as ex:
        log(f"临时文件清理异常（不影响任务）：{ex}")
        return 0


def _task_worker():
    """唯一的任务执行线程：从队列按入队顺序取，天然串行

    队列元素 (mode, day, payload)：normal=常规单日，recalc=周期末日整周期重算，
    fix=修正版范围重算。三者可能同时排队，必须串行且升序，因为后一天依赖前一天。
    """
    while True:
        mode, day, payload = _RUN_QUEUE.get()
        before = snapshot_scripts()
        try:
            if mode == "fix":
                run_fix(payload)
            elif mode == "recalc":
                run_recalc(day)
            else:
                run_task(day)
        except Exception as e:
            log(f"任务异常 {day}｜{e}")
        finally:
            try:
                sweep_scripts(day, before)
            except Exception as e:
                log(f"临时文件收纳异常 {day}｜{e}")
            purge_tmp()          # 内部自带节流与全量保底，出错只记日志
            _QUEUED.discard((mode, day))
            _RUN_QUEUE.task_done()


def enqueue_task(day: str, mode: str = "normal", payload=None) -> bool:
    """入队（去重）。调用方按日期升序调用，队列保证按此顺序执行"""
    key = (mode, day)
    if key in _QUEUED:
        return False
    _QUEUED.add(key)
    _RUN_QUEUE.put((mode, day, payload))
    return True


def _is_queued(day: str, mode: str = None) -> bool:
    if mode is not None:
        return (mode, day) in _QUEUED
    return any((m, day) in _QUEUED for m in ("normal", "recalc", "fix"))


def run_task(day: str):
    """跑一次任务：脚本优先，模型兜底。同日跑过就跳过。

    为什么先试脚本（第七版）：脚本秒级、零 token、同输入同输出，出错能复现；
    模型适合处理「口径不明」这种真需要判断的事，不适合当计算器。分流规则只有
    一条 —— 脚本明确说「口径不明」(exit 2) 才转模型，其余终态一律就地解决：
      0/4 有产物 → 完成；0/4 无产物 → 假完成，记失败
      3 自检不过 → 拦在暂存区，**不能**转模型（模型一看「已有产物」就当完成，错数就此固化）
      5 缺件     → 转模型也造不出（口径禁止推测值）
      None 环境问题（解释器/脚本缺失、超时）→ 转模型同样跑不动，只会白烧额度
    """
    done = TRIGGER / "_done" / f"{day}.ok"
    lock = TRIGGER / "_running" / f"{day}.lock"
    done.parent.mkdir(parents=True, exist_ok=True)
    lock.parent.mkdir(parents=True, exist_ok=True)

    if done.exists():
        log(f"跳过：{day} 已完成")
        return
    if lock.exists():
        # 断电/进程被杀时锁不会被清掉，会把该日期永久堵死，
        # 所以超过「单次超时 + 5 分钟」的锁一律视为残骸，清掉继续。
        try:
            age = time.time() - lock.stat().st_mtime
        except OSError:
            age = 0.0
        if age > CFG["run_timeout"] + 300:
            log(f"清理陈旧执行锁：{day}（残留 {age / 60:.0f} 分钟，判定为上次异常退出）")
            lock.unlink(missing_ok=True)
        else:
            log(f"跳过：{day} 正在执行中")
            return

    out_log = LOGDIR / f"run-{day}.log"
    lock.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
    t0 = time.time()
    try:
        # ---------------------------------------------------------- 第一路：确定性脚本
        if CFG.get("script_first") and (CFG.get("script_cmd") or []):
            slog = LOGDIR / f"script-{day}.log"
            log(f"脚本通道开始：{day}")
            rc, text, why = run_script(day, slog, float(CFG.get("script_timeout", 900)))
            elapsed = time.time() - t0

            if rc in (0, 4) and has_produced(day):
                _finish_done(day, done, elapsed, "script" if rc == 0 else "script-already")
                conc = CONC_DIR / f"{day}.md"
                if not conc.exists():
                    conc = save_script_conclusion(
                        day, rc, text, elapsed,
                        "成功" if rc == 0 else "当日此前已完成")
                log(f"脚本通道完成：{day}（exit {rc}，{elapsed:.1f} 秒）｜结论见 {conc.name}")
                notify("任务完成",
                       f"{day} 脚本通道完成，用时 {elapsed:.1f} 秒。\n"
                       f"结论：{CONC_DIR.name}/{conc.name}",
                       sticky=True, open_path=CONC_DIR)
                return

            if rc is None:
                n, wait = _mark_fail(day, why or "脚本未能启动")
                save_script_conclusion(day, rc, text or why, elapsed, "脚本未能启动")
                log(f"脚本通道无法启动：{day}｜{why}｜第 {n} 次失败，"
                    f"{wait // 60} 分钟后重试")
                notify("脚本通道异常",
                       f"{day} 没能启动生成脚本。\n原因：{why}\n"
                       f"第 {n} 次失败，约 {wait // 60} 分钟后自动重试。",
                       sticky=True, open_path=slog)
                return

            if rc in (0, 4) and not has_produced(day):
                why = f"脚本 exit {rc} 但目录里没有产物"
                n, wait = _mark_fail(day, why)
                save_script_conclusion(day, rc, text, elapsed, "脚本报完成但未产出")
                log(f"脚本通道异常：{day}｜{why}｜第 {n} 次失败，"
                    f"{wait // 60} 分钟后重试｜详见 {slog.name}")
                notify("任务未产出",
                       f"{day} 脚本报完成但目录里找不到产物。\n"
                       f"第 {n} 次失败，约 {wait // 60} 分钟后自动重试。\n查看 {slog}",
                       sticky=True, open_path=slog)
                return

            if rc == 5:
                why = _last_line(text)
                n, wait = _mark_fail(day, why)
                save_script_conclusion(day, rc, text, elapsed, "缺件未执行")
                log(f"脚本通道缺件：{day}｜{why}｜第 {n} 次，"
                    f"{wait // 60} 分钟后重试｜详见 {slog.name}")
                notify("缺件未执行",
                       f"{day} 源件不齐，脚本按口径停下（不猜、不替代）。\n"
                       f"{why}\n第 {n} 次，约 {wait // 60} 分钟后自动重试。",
                       sticky=True, open_path=slog)
                return

            if rc == 3:
                why = _last_line(text)
                n, wait = _mark_fail(day, "自检不过：" + why)
                save_script_conclusion(day, rc, text, elapsed, "自检不过（未落盘）")
                log(f"脚本通道自检不过：{day}｜{why}｜第 {n} 次，"
                    f"{wait // 60} 分钟后重试｜详见 {slog.name}")
                notify("自检不过（未落盘）",
                       f"{day} 生成结果没通过自检，已拦在暂存区，正式产物未被替换。\n"
                       f"{why}\n完整自检结果见 {slog.name}。\n"
                       f"第 {n} 次，约 {wait // 60} 分钟后自动重试。",
                       sticky=True, open_path=slog)
                return

            # exit 2 口径不明：唯一需要判断的路径，交模型
            log(f"脚本通道未通过（exit {rc}，{elapsed:.1f} 秒）：{_last_line(text)}"
                f"｜转入模型兜底")

        # ---------------------------------------------------------- 第二路：模型兜底
        log(f"开始执行：{day}｜输出 {out_log.name}")
        r = _run_cli_retry(build_prompt(day), out_log, CFG["run_timeout"],
                           f"执行 {day}", extra_args=_perm_args())
        elapsed = time.time() - t0
        if r.returncode == 0 and has_produced(day):
            _finish_done(day, done, elapsed, "cli")
            log(f"执行完成：{day}（exit 0，耗时 {elapsed / 60:.1f} 分钟）")
            save_conclusion(day, r, elapsed, out_log)
            notify("任务完成", f"{day} 已执行，结论见 {CONC_DIR.name}/{day}.md",
                   sticky=True, open_path=CONC_DIR)
        elif r.returncode == 0:
            # 正常退出却没产物：绝不能写 done，否则 _day_action 按「已完成」跳过，
            # 该日永久漏掉，等到周期末日重算才发现少一天的 sheet。
            n, wait = _mark_fail(day, "模型正常退出但未产出")
            log(f"执行未产出：{day}（exit 0，{elapsed / 60:.1f} 分钟），"
                f"第 {n} 次失败，{wait // 60} 分钟后重试｜详见 {out_log.name}")
            save_conclusion(day, r, elapsed, out_log)
            notify("任务未产出",
                   f"{day} 命令正常退出但未生成产物。\n"
                   f"第 {n} 次失败，约 {wait // 60} 分钟后自动重试。\n查看 {out_log}",
                   sticky=True, open_path=out_log)
        else:
            n, wait = _mark_fail(day, f"exit {r.returncode}")
            log(f"执行失败：{day}（exit {r.returncode}），第 {n} 次失败，"
                f"{wait // 60} 分钟后重试｜详见 {out_log.name}")
            save_conclusion(day, r, elapsed, out_log)
            notify("任务失败", f"{day} exit {r.returncode}，第 {n} 次失败。\n"
                            f"约 {wait // 60} 分钟后自动重试。",
                   sticky=True, open_path=out_log)
    except Exception as e:
        log(f"执行异常：{day}｜{e}")
    finally:
        lock.unlink(missing_ok=True)


def _finish_done(day: str, done: Path, elapsed: float, via: str):
    """写完成标记并清掉失败计数。脚本路与模型路共用。"""
    done.write_text(json.dumps({
        "day": day,
        "at": datetime.now().isoformat(timespec="seconds"),
        "elapsed_seconds": round(elapsed),
        "exit": 0,
        "via": via,
    }, ensure_ascii=False), encoding="utf-8")
    _clear_fail(day)


def run_task_serial(day: str):
    """按日串行（保留给手动调用；主循环请用 enqueue_task）"""
    if not _RUN_SERIAL.acquire(timeout=CFG["run_timeout"] + 600):
        log(f"排队超时：{day}")
        return
    try:
        run_task(day)
    finally:
        _RUN_SERIAL.release()


# ---------------------------------------------------------------- 报表周期与重算
# 两种「产物已经出来、但结果是错的」的场景：
#   1) 周期末日：业务方在周期最后一天重发源文件并改了口径（典型是单重/系数），
#      本周期每一天的数字都变 → 必须整周期重算，只补末日那一天等于留一个月的错数。
#   2) 任一日期被修正版替换：这份源到底改了哪几天，外部看不出来，只能按「源的覆盖范围」
#      推定影响面 —— 累计堆叠型（coverage=cumulative）从周期首日算，单日型从被替换那天算，
#      终点都取「当期最新已产出」，没出的日子交给常规流程。
# 两个场景都依赖 produced_glob（否则无从判断产物到哪天为止），未配置则自动不启用。
def _cycle_cfg():
    c = CFG.get("cycle") or {}
    return {
        "enabled": bool(c.get("enabled")),
        "start_day": int(c.get("start_day", 26)),
        "end_day": int(c.get("end_day", 25)),
        "recalc_on_end": bool(c.get("recalc_on_end", True)),
        "timeout": float(c.get("timeout") or CFG.get("run_timeout", 5400)),
    }


def _shift_month(y, m, k):
    idx = y * 12 + (m - 1) + k
    return idx // 12, idx % 12 + 1


def period_bounds(day: str):
    """该日所属周期 (首日, 末日)。默认 26 日 ~ 次月 25 日"""
    dt = datetime.fromisoformat(day).date()
    sd, ed = _cycle_cfg()["start_day"], _cycle_cfg()["end_day"]
    if dt.day >= sd:
        y2, m2 = _shift_month(dt.year, dt.month, 1)
        s, e = dt.replace(day=sd), dt.replace(year=y2, month=m2, day=ed)
    else:
        y0, m0 = _shift_month(dt.year, dt.month, -1)
        s, e = dt.replace(year=y0, month=m0, day=sd), dt.replace(day=ed)
    if e < s:
        e = dt
    return s, e


def period_days(any_day: str):
    """周期内全部日期，升序 ISO 字符串"""
    s, e = period_bounds(any_day)
    out, d = [], s
    while d <= e:
        out.append(d.isoformat())
        d += timedelta(days=1)
    return out


def is_period_end(day: str) -> bool:
    c = _cycle_cfg()
    return c["enabled"] and int(day.split("-")[2]) == c["end_day"]


def latest_output_day(day: str):
    """当期已产出产物的最晚日期；查不到返回 None"""
    today = datetime.now().date().isoformat()
    for d in reversed(period_days(day)):
        if d > today:
            continue
        if has_produced(d):
            return d
    return None


def _day_action(day: str, done_dir: Path):
    """返回 (action, reason)。action: skip / recalc / normal

    周期末日即使已有产物、已有完成标记，只要没有 recalc 标记就必须重算 ——
    那一天源文件里改的口径会让全周期数字都失效，不能按「已产出」跳过。
    """
    c = _cycle_cfg()
    if c["enabled"] and c["recalc_on_end"] and is_period_end(day):
        if (done_dir / f"recalc-{day}.ok").exists():
            return "skip", "整周期重算已完成"
        return "recalc", "周期末日需整周期重算"
    if (done_dir / f"{day}.ok").exists():
        return "skip", "已完成"
    if has_produced(day):
        return "skip", "已有产物"
    return "normal", ""


def _scope_addendum(start: str, end: str, full: bool, why: str) -> str:
    """追加到用户配置的 prompt 之后。模板不知道你的产物长什么样（一文件多 sheet、
    一天一文件都可能），所以这里只写与结构无关的硬要求。"""
    s, e = period_bounds(end)
    n = len([d for d in period_days(end) if start <= d <= end])
    prev = (datetime.fromisoformat(start).date() - timedelta(days=1))
    base = ("" if full else
            f"；{start} 不是周期首日，它的累计基数取上一日（{prev.month}.{prev.day}）的现值，"
            f"那一日不在本次范围内，只读不改")
    return (
        f"\n\n【本轮是范围重算，不是常规单日出报】\n"
        f"报表周期：{s.isoformat()} 至 {e.isoformat()}。\n"
        f"本次重算范围：{start} 至 {end}，共 {n} 天"
        + ("（本周期全部日期）" if full else "") + "。\n"
        f"原因：{why}\n"
        f"要求：\n"
        f"1) 按日期**升序**逐日重算，每天的数字都要回到源文件重新取，"
        f"**禁止沿用产物里的旧值**。\n"
        f"2) 范围外的日期一律不动（{start} 之前、{end} 之后都不碰）。\n"
        f"3) 累计/结转类数字必须引用刚重算过的上一日，不得引用旧快照{base}。\n"
        f"4) 源文件缺失的日期：保留原值并在摘要中逐日列出，禁止用推测值或邻日值替代。\n"
        f"5) 自检：区间末日的累计值应等于逐日数值之和；出现倒退或断档如实报告，不要硬凑。\n"
        f"6) 结束时用纯文本输出：重算了哪些日期、区间末日的累计值、与重算前相比变化多少、"
        f"缺源的日期清单。不生成额外文件。"
    )


def run_scope_task(start: str, end: str, full: bool, why: str, tag: str, timeout: float):
    """按范围重算 [start, end]。tag 同时用作完成标记、锁、日志与结论文件名。"""
    done = TRIGGER / "_done" / f"{tag}.ok"
    lock = TRIGGER / "_running" / f"{tag}.lock"
    done.parent.mkdir(parents=True, exist_ok=True)
    lock.parent.mkdir(parents=True, exist_ok=True)

    if done.exists():
        log(f"跳过重算：{tag} 已完成")
        return
    if lock.exists():
        try:
            age = time.time() - lock.stat().st_mtime
        except OSError:
            age = 0.0
        if age > timeout + 300:
            log(f"清理陈旧重算锁：{tag}（残留 {age / 60:.0f} 分钟，判定为上次异常退出）")
            lock.unlink(missing_ok=True)
        else:
            log(f"跳过重算：{tag} 正在执行中")
            return

    out_log = LOGDIR / f"{tag}.log"
    lock.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
    t0 = time.time()
    try:
        log(f"开始范围重算：{tag}｜{start}~{end}"
            + ("（全周期）" if full else "") + f"｜输出 {out_log.name}")
        notify("开始范围重算", f"{tag}：{start} 至 {end}")
        prompt = build_prompt(end) + _scope_addendum(start, end, full, why)
        r = _run_cli_retry(prompt, out_log, timeout, tag, extra_args=_perm_args())
        elapsed = time.time() - t0
        save_conclusion(tag, r, elapsed, out_log)
        if r.returncode == 0:
            done.write_text(json.dumps(
                {"tag": tag, "scope": [start, end], "full": bool(full),
                 "at": datetime.now().isoformat(timespec="seconds"),
                 "elapsed_seconds": round(elapsed, 1), "exit": 0},
                ensure_ascii=False), encoding="utf-8")
            log(f"范围重算完成：{tag}（exit 0，{elapsed / 60:.1f} 分钟）")
            notify("范围重算完成", f"{tag}：{start} 至 {end}，见 {CONC_DIR.name}/{tag}.md",
                   sticky=True, open_path=CONC_DIR)
        else:
            log(f"范围重算失败：{tag}（exit {r.returncode}）｜详见 {out_log.name}")
            notify("范围重算失败", f"{tag} exit {r.returncode}", sticky=True, open_path=out_log)
    except Exception as e:
        log(f"范围重算异常：{tag}｜{e}")
    finally:
        lock.unlink(missing_ok=True)


# ---- 修正版重算：源文件被替换 → 按影响面重算一段区间
FIX_PENDING = TRIGGER / "_fix_pending.json"
FIX_STATE = TRIGGER / "_fix_state.json"
SRC_SNAP = TRIGGER / "_source_fps.json"


def _fix_cfg():
    c = CFG.get("fix_recalc") or {}
    return {
        "enabled": bool(c.get("enabled", True)),
        "cumulative_full_period": bool(c.get("cumulative_full_period", True)),
        "quiet_minutes": float(c.get("quiet_minutes", 10)),
        "max_per_day": int(c.get("max_per_day", 3)),
        "allow_closed_period": bool(c.get("allow_closed_period", False)),
        "timeout": float(c.get("timeout") or CFG.get("run_timeout", 5400)),
    }


def _read_json(p: Path, default):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return default


def _write_json(p: Path, obj):
    try:
        Path(p).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        log(f"写 {Path(p).name} 失败：{e}")


def _coverage_of(name: str) -> str:
    for s in SOURCES:
        if s.name == name:
            return s.coverage
    return "single"


def _source_fps():
    """各「源+日期」的版本指纹，只 stat 不读内容，只扫补日窗口内的日期"""
    look = max(1, int(CFG.get("lookback_days", 7)))
    today = datetime.now().date()
    window = {(today - timedelta(days=i)).isoformat() for i in range(look)}
    out = {}
    for s in SOURCES:
        for d, f in s.list_all():
            ds = d.isoformat()
            if ds not in window:
                continue
            try:
                st = f.stat()
            except OSError:
                continue
            out[f"{s.name}|{ds}"] = {"fp": f"{st.st_size}:{int(st.st_mtime)}", "path": str(f)}
    return out


def note_fix_event(category: str, day: str, dst, why: str = ""):
    """登记一次「源文件被新版替换」。先进待办，静默期过后再收 —— 对方连发两版只跑最后那版。"""
    fc = _fix_cfg()
    if not fc["enabled"] or not category or not day:
        return
    if is_period_end(day) and _cycle_cfg()["recalc_on_end"]:
        log(f"修正版并入整周期重算：{category} {day}")
        return
    if not has_produced(day):
        log(f"修正版忽略：{category} {day} 尚未产出，常规流程会用最新源")
        return
    try:
        st = Path(dst).stat()
        fp = f"{st.st_size}:{int(st.st_mtime)}"
    except OSError:
        return
    pend = _read_json(FIX_PENDING, {})
    key = f"{category}|{day}"
    now = time.time()
    it = pend.get(key) or {"category": category, "day": day,
                           "first_seen": now, "versions": 0}
    it.update({"dst": str(dst), "fp": fp, "last_seen": now,
               "versions": int(it.get("versions", 0)) + 1})
    pend[key] = it
    _write_json(FIX_PENDING, pend)
    log(f"登记修正版：{category} {day}｜第 {it['versions']} 版｜"
        f"静默 {fc['quiet_minutes']:.0f} 分钟无新版本后重算" + (f"｜{why}" if why else ""))


def watch_source_changes():
    """扫一遍源文件版本，和上次快照比：某「源+日期」变过且该日已产出 → 登记修正版重算。

    走快照对比而不是只挂在「外部取件」这一步，是因为使用者也可能直接覆盖源目录里的
    文件（或手工另存），那条路径不经过 adopt()，只在取件处挂钩会漏。
    """
    cur = _source_fps()
    snap = _read_json(SRC_SNAP, {})
    for key, info in cur.items():
        old = (snap.get(key) or {}).get("fp")
        if old and old != info["fp"]:
            name, day = key.split("|", 1)
            note_fix_event(name, day, Path(info["path"]), "源文件版本变化")
    _write_json(SRC_SNAP, cur)


def fix_scope(category: str, day: str):
    """这份修正版影响到哪些日期。返回 (scope, reason)，scope=(start, end, full)"""
    fc = _fix_cfg()
    s, e = period_bounds(day)
    if is_period_end(day) and _cycle_cfg()["recalc_on_end"]:
        return None, "周期末日归整周期重算统一处理"
    if e < datetime.now().date() and not fc["allow_closed_period"]:
        return None, (f"该日所属周期（{s.isoformat()}~{e.isoformat()}）已结束，"
                      f"不自动重算历史周期，请人工确认")
    if not has_produced(day):
        return None, "该日尚未产出，常规流程会用最新源，无需补算"
    latest = latest_output_day(day)
    if not latest:
        return None, "当期查不到已产出的产物（produced_glob 未配置或尚未产出）"
    if day > latest:
        return None, f"该日（{day}）晚于当期最新产物（{latest}）"
    full = (_coverage_of(category) == "cumulative" and fc["cumulative_full_period"])
    start = s.isoformat() if full else day
    if start > latest:
        return None, "区间起点晚于最新产物"
    return (start, latest, full), ""


def _fix_runs_today(key: str) -> int:
    st = _read_json(FIX_STATE, {})
    today = datetime.now().date().isoformat()
    return sum(1 for x in ((st.get(key) or {}).get("runs") or [])
               if str(x.get("at", "")).startswith(today))


def _mark_fix(key: str, fp: str, scope=None, fail=False):
    st = _read_json(FIX_STATE, {})
    it = st.setdefault(key, {})
    if not fail and scope:
        it["fp"] = fp
        it["scope"] = list(scope)
    it["applied_at"] = datetime.now().isoformat(timespec="seconds")
    runs = it.get("runs") or []
    runs.append({"fp": fp, "at": it["applied_at"], "fail": bool(fail)})
    it["runs"] = runs[-20:]
    _write_json(FIX_STATE, st)


def collect_ready_fixes():
    """收取可以跑的修正版重算：静默期已过、版本没跑过、次数没超、区间有效。"""
    fc = _fix_cfg()
    if not fc["enabled"]:
        return []
    pend = _read_json(FIX_PENDING, {})
    if not pend:
        return []
    st = _read_json(FIX_STATE, {})
    now, out, dirty = time.time(), [], False
    for key in list(pend):
        it = pend.get(key) or {}
        cat, day = it.get("category"), it.get("day")
        if not cat or not day:
            pend.pop(key, None)
            dirty = True
            continue
        fp = str(it.get("fp") or "")
        if fp and (st.get(key) or {}).get("fp") == fp:
            pend.pop(key, None)
            dirty = True
            log(f"修正版待办作废：{cat} {day} 该版本已重算过")
            continue
        if now - float(it.get("last_seen") or 0) < fc["quiet_minutes"] * 60:
            continue
        scope, why = fix_scope(cat, day)
        if not scope:
            pend.pop(key, None)
            dirty = True
            log(f"修正版待办作废：{cat} {day}｜{why}")
            continue
        n = _fix_runs_today(key)
        if n >= fc["max_per_day"]:
            pend.pop(key, None)
            dirty = True
            log(f"修正版待办作废：{cat} {day} 今日已重算 {n} 次，达上限 {fc['max_per_day']}")
            notify("修正版已达次数上限", f"{cat} {day} 今日已重算 {n} 次，请人工确认",
                   sticky=True, open_path=CONC_DIR)
            continue
        start, end, full = scope
        out.append({"start": start, "end": end, "full": full, "why": why,
                    "tag": f"fix-{day}-{cat}-{fp.replace(':', '-')}",
                    "key": key, "day": day, "cat": cat, "fp": fp})
        pend.pop(key, None)
        dirty = True
    if dirty:
        _write_json(FIX_PENDING, pend)
    return sorted(out, key=lambda x: (x["day"], x["cat"]))


def run_fix(fx: dict):
    """跑一次修正版重算，成功/失败都记次数（失败也记，避免坏任务无限重试）"""
    tag, key, fp = fx["tag"], fx["key"], fx["fp"]
    try:
        run_scope_task(fx["start"], fx["end"], fx["full"], fx["why"], tag,
                       _fix_cfg()["timeout"])
        ok = (TRIGGER / "_done" / f"{tag}.ok").exists()
        _mark_fix(key, fp, [fx["start"], fx["end"]], fail=not ok)
        if not ok:
            pend = _read_json(FIX_PENDING, {})
            pend[key] = {"category": fx["cat"], "day": fx["day"],
                         "dst": "", "fp": fp,
                         "first_seen": time.time(), "last_seen": time.time(),
                         "versions": 1}
            _write_json(FIX_PENDING, pend)
            log(f"修正版重算未成功：{tag}，已重新登记待下次静默期后重试")
    except Exception as e:
        log(f"修正版重算异常：{tag}｜{e}")


def run_recalc(end_day: str):
    """周期末日整周期重算"""
    s, e = period_bounds(end_day)
    start = s.isoformat()
    why = (f"业务方在周期末日重发的源文件里改了口径（常见是单重/系数），"
           f"由「数量 × 系数」算出的数在**本周期每一天**都变了；"
           f"累计列是链式的，某天变了之后每一天都得跟着变。")
    run_scope_task(start, e.isoformat(), True, why, f"recalc-{end_day}",
                   _cycle_cfg()["timeout"])


# ---------------------------------------------------------------- 单实例
def _pid_alive(pid: int) -> bool:
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = ctypes.c_ulong()
        ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
        k32.CloseHandle(h)
        return bool(ok) and code.value == 259
    except Exception:
        return False


def acquire_single_instance():
    lock = TRIGGER / "_watcher.lock"
    if lock.exists():
        try:
            old = int(lock.read_text(encoding="utf-8").strip() or "0")
        except Exception:
            old = 0
        if old and old != os.getpid() and _pid_alive(old):
            return None
        lock.unlink(missing_ok=True)
    lock.write_text(str(os.getpid()), encoding="utf-8")
    return lock


# ---------------------------------------------------------------- 主循环
def main():
    lock = acquire_single_instance()
    if lock is None:
        log("已在运行，本次启动退出（单实例保护）")
        return
    if not SOURCES:
        log("未配置任何 sources，退出")
        return
    log(f"启动｜根目录 {ROOT}｜源 {[s.name for s in SOURCES]}｜轮询 {CFG['poll_seconds']}s｜pid {os.getpid()}")
    _beat(force=True)

    today = datetime.now().date().isoformat()
    mark = TRIGGER / "_preflight" / f"{today}.ok"
    mark.parent.mkdir(parents=True, exist_ok=True)
    if CFG.get("preflight_on_start") and not (mark.exists() and (TRIGGER / "_perm_mode.json").exists()):
        if not check_ready(today)[0]:
            mark.write_text(datetime.now().isoformat(), encoding="utf-8")
            threading.Thread(target=run_preflight, daemon=True).start()

    threading.Thread(target=_task_worker, daemon=True).start()
    log(f"任务执行线程已启动｜补日窗口 {CFG.get('lookback_days', 7)} 天")

    ext = _ext_cfgs()
    ext_next = {c["name"]: 0.0 for c in ext}
    if ext:
        log("外部来源已启用｜" + "；".join(f"{c['name']} <- {c['root']}" for c in ext))

    warned = set()
    try:
        while True:
            try:
                _beat()
                # 弹窗链路自检：写 <trigger>/_popup_probe.txt 即弹（第九版）
                _check_popup_probe()
                # 外部来源取件：文件先落在下载目录/聊天软件目录时，靠这一步认领
                for c in ext:
                    if time.time() < ext_next[c["name"]]:
                        continue
                    ext_next[c["name"]] = time.time() + c["interval_seconds"]
                    try:
                        pickup_external(dedupe_external(scan_external(c)))
                    except Exception as e:
                        log(f"外部取件异常 {c['name']}：{e}")

                # 源文件版本巡检：某「源+日期」被新版替换且该日已出产物 → 登记修正版重算
                try:
                    watch_source_changes()
                except Exception as e:
                    log(f"源版本巡检异常：{e}")

                # 修正版重算排在「齐备检查」之前入队：先把旧日期修正，再出新日期的产物，
                # 否则新产物会踩着尚未修正的上一日累计值算。
                for fx in collect_ready_fixes():
                    if _is_queued(fx["day"], "fix"):
                        continue
                    desc = f"{fx['cat']} {fx['day']} 修正版｜重算 {fx['start']}~{fx['end']}"
                    if CFG.get("auto_run"):
                        if enqueue_task(fx["day"], "fix", fx):
                            log(f"已入队修正版重算：{desc}（队列待执行 {_RUN_QUEUE.qsize()}）")
                    else:
                        notify("源文件出现修正版", desc + "（auto_run=false，需人工触发）")

                # 齐备检查：最近 N 天内、尚未出产物、尚未开跑的日期。
                # 源文件到位时间不固定（可能今天才到前天的、也可能今天的明天才到），
                # 所以窗口由 lookback_days 决定；已出过产物的一律跳过，绝不动已有结果。
                today = datetime.now().date()
                look = max(1, int(CFG.get("lookback_days", 7)))
                days = [(today - timedelta(days=i)).isoformat() for i in range(look)]
                done_dir = TRIGGER / "_done"
                for day in sorted(days):
                    if _is_queued(day):
                        continue
                    action, reason = _day_action(day, done_dir)
                    if action == "skip":
                        if day not in warned:
                            warned.add(day)
                            log(f"{day} {reason}，跳过（被新版替换的情况走修正版重算）")
                        continue
                    ok, missing, rp = check_ready(day)
                    if ok:
                        _MISS.pop(day, None)
                        if day not in warned:
                            warned.add(day)
                            if action == "recalc":
                                s0, e0 = period_bounds(day)
                                log(f"源文件齐备（{day}，周期末日）｜"
                                    f"{s0.isoformat()}~{e0.isoformat()} 共 "
                                    f"{len(period_days(day))} 天，将整周期重算")
                            else:
                                log(f"源文件齐备（{day}），标记 {rp.name}")
                            if CFG.get("auto_run"):
                                # 退避闸：被记过失败的日子，没到重试时间就不入队（第八版）。
                                # 这一步必须用落盘状态判断，不能用内存集合 —— 否则某日
                                # 入队失败一次就永久卡死，日志却一直写「下一轮重试」。
                                ok_run, attempt, last_err = _retry_gate(day)
                                if not ok_run:
                                    if attempt and last_err:
                                        log(f"{day} 退避中（已试 {attempt} 次，"
                                            f"上次：{last_err[:80]}）")
                                    continue
                                if enqueue_task(day, action):
                                    log(f"已入队：{day}"
                                        f"（{'整周期重算' if action == 'recalc' else '常规'}，"
                                        f"队列待执行 {_RUN_QUEUE.qsize()}"
                                        f"{'，第 ' + str(attempt) + ' 次' if attempt > 1 else ''}）")
                            else:
                                notify("源文件已齐备",
                                       f"{day} 可执行任务（auto_run=false，需人工触发）")
                    else:
                        warned.discard(day)
                        # 缺件状态变化时记一次；之后按心跳间隔留一行，证明守着没死
                        sig = tuple(missing)
                        prev = _MISS.get(day)
                        now_ts = time.time()
                        hb = max(60, int(CFG.get("wait_log_interval_seconds", 1800)))
                        if prev is None:
                            _MISS[day] = (sig, now_ts)
                            log(f"{day} 等待源文件：缺 {'、'.join(missing)}")
                        elif prev[0] != sig:
                            _MISS[day] = (sig, now_ts)
                            log(f"{day} 缺件有变化：现缺 {'、'.join(missing) if missing else '（写入中）'}")
                        elif now_ts - prev[1] >= hb:
                            _MISS[day] = (sig, now_ts)
                            log(f"{day} 仍在等待源文件：缺 {'、'.join(missing)}"
                                f"（监听正常，每 {hb // 60} 分钟报一次）")
                time.sleep(CFG["poll_seconds"])
            except Exception as e:
                log(f"主循环异常：{e}")
                time.sleep(10)
    except KeyboardInterrupt:
        log("停止")
    finally:
        try:
            lock.unlink(missing_ok=True)
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        for c in _ext_cfgs():
            n = pickup_external(dedupe_external(scan_external(c)))
            log(f"外部取件 [{c['name']}]：搬运 {n} 个")
        day = sys.argv[2] if len(sys.argv) > 2 else datetime.now().date().isoformat()
        ok, missing, rp = check_ready(day)
        log(f"单次检查 {day}：{'齐备' if ok else '缺 ' + ','.join(missing)}")
        print(build_prompt(day))
        sys.exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == "--pickup":
        for c in _ext_cfgs():
            cands = dedupe_external(scan_external(c))
            log(f"外部来源 [{c['name']}]：候选 {len(cands)} 个")
            for s, p in cands:
                log(f"  {s.name}｜{p.name}｜{p.stat().st_size} B")
            n = pickup_external(cands)
            log(f"实际搬运 {n} 个")
        sys.exit(0)
    main()
