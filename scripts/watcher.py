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
for d in (LOGDIR, TRIGGER, CONC_DIR, STREAM_DIR):
    d.mkdir(parents=True, exist_ok=True)

IGNORE_PREFIX = ("~$", "._", ".")
# 判定「谁更新」时容许的 mtime 误差，避免复制动作本身带来的毫秒级抖动误判
MTIME_TOL = 2.0


# ---------------------------------------------------------------- 工具
def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    try:
        with open(LOGDIR / f"watcher-{datetime.now():%Y-%m}.log", "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass


def notify(title, text):
    """任务完成提示。默认关声音，需要时在配置里开。"""
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass
    log(f"[通知] {title}｜{text}")


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


def _task_worker():
    """唯一的任务执行线程：从队列按入队顺序取，天然串行"""
    while True:
        day = _RUN_QUEUE.get()
        try:
            run_task(day)
        except Exception as e:
            log(f"任务异常 {day}｜{e}")
        finally:
            _QUEUED.discard(day)
            _RUN_QUEUE.task_done()


def enqueue_task(day: str) -> bool:
    """入队（去重）。调用方按日期升序调用，队列保证按此顺序执行"""
    if day in _QUEUED:
        return False
    _QUEUED.add(day)
    _RUN_QUEUE.put(day)
    return True


def run_task(day: str):
    """跑一次任务。同日跑过就跳过；成功的标记写 <trigger>/_done/<day>.ok"""
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
        log(f"开始执行：{day}｜输出 {out_log.name}")
        r = _run_cli_retry(build_prompt(day), out_log, CFG["run_timeout"],
                           f"执行 {day}", extra_args=_perm_args())
        elapsed = time.time() - t0
        if r.returncode == 0:
            done.write_text(json.dumps(
                {"day": day, "at": datetime.now().isoformat(timespec="seconds"),
                 "elapsed_seconds": round(elapsed, 1), "exit": 0},
                ensure_ascii=False), encoding="utf-8")
            log(f"执行完成：{day}（exit 0，耗时 {elapsed / 60:.1f} 分钟）")
            save_conclusion(day, r, elapsed, out_log)
            notify("任务完成", f"{day} 已执行，结论见 {CONC_DIR.name}/{day}.md")
        else:
            log(f"执行失败：{day}（exit {r.returncode}）｜详见 {out_log.name}")
            save_conclusion(day, r, elapsed, out_log)
            notify("任务失败", f"{day} exit {r.returncode}")
    except Exception as e:
        log(f"执行异常：{day}｜{e}")
    finally:
        lock.unlink(missing_ok=True)


def run_task_serial(day: str):
    """按日串行（保留给手动调用；主循环请用 enqueue_task）"""
    if not _RUN_SERIAL.acquire(timeout=CFG["run_timeout"] + 600):
        log(f"排队超时：{day}")
        return
    try:
        run_task(day)
    finally:
        _RUN_SERIAL.release()


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
                # 外部来源取件：文件先落在下载目录/聊天软件目录时，靠这一步认领
                for c in ext:
                    if time.time() < ext_next[c["name"]]:
                        continue
                    ext_next[c["name"]] = time.time() + c["interval_seconds"]
                    try:
                        pickup_external(dedupe_external(scan_external(c)))
                    except Exception as e:
                        log(f"外部取件异常 {c['name']}：{e}")

                # 齐备检查：最近 N 天内、尚未出产物、尚未开跑的日期。
                # 源文件到位时间不固定（可能今天才到前天的、也可能今天的明天才到），
                # 所以窗口由 lookback_days 决定；已出过产物的一律跳过，绝不动已有结果。
                today = datetime.now().date()
                look = max(1, int(CFG.get("lookback_days", 7)))
                days = [(today - timedelta(days=i)).isoformat() for i in range(look)]
                done_dir = TRIGGER / "_done"
                for day in sorted(days):
                    if day in _QUEUED:
                        continue
                    if (done_dir / f"{day}.ok").exists():
                        continue
                    if has_produced(day):
                        if day not in warned:
                            warned.add(day)
                            log(f"{day} 已有产物，跳过（不重算）")
                        continue
                    ok, missing, rp = check_ready(day)
                    if ok:
                        _MISS.pop(day, None)
                        if day not in warned:
                            warned.add(day)
                            log(f"源文件齐备（{day}），标记 {rp.name}")
                            if CFG.get("auto_run"):
                                if enqueue_task(day):
                                    log(f"已入队：{day}（队列待执行 {_RUN_QUEUE.qsize()}）")
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
