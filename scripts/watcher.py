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
  - 补日场景会并发触发多天任务，必须串行排队，否则同写一个输出文件会互相覆盖。

目录约定：
  <project>/scripts/watcher.py     ← 本文件
  <project>/scripts/config.json    ← 配置（可复制 config.example.json 改名）
  <project>/<source dir>/...       ← 各源文件目录
  <project>/<log_dir>/             ← 运行日志（自动创建）
  <project>/<trigger_dir>/         ← ready 标记、锁（自动创建）
"""

import os
import re
import sys
import json
import time
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
    "poll_seconds": 5,
    "stable_seconds": 3,
    "lookback_days": 1,
    "auto_run": True,
    "preflight_on_start": True,
    "run_timeout": 5400,
    "sources": [],
    "task": {"spec_doc": "", "prompt": ""},
    "cli": {
        "node": "",
        "entry": "",
        "git_bash": "",
        "product_version": "",
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
for d in (LOGDIR, TRIGGER):
    d.mkdir(parents=True, exist_ok=True)

IGNORE_PREFIX = ("~$", "._", ".")


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


# ---------------------------------------------------------------- 源文件解析
class Source:
    """一个源：目录 + 文件名正则（用命名分组 y/m/d 取日期）"""

    def __init__(self, spec):
        self.name = spec.get("name") or spec.get("dir", "source")
        self.dir = ROOT / spec["dir"]
        self.pattern = re.compile(spec["pattern"])
        self.allow_newer = bool(spec.get("allow_newer_fallback", False))
        self.std_name = spec.get("std_name", "")  # 可选：如 "A_{m}-{d}.xlsx"

    def std_path(self, day: str) -> Path:
        y, m, d = [int(x) for x in day.split("-")]
        if self.std_name:
            return self.dir / self.std_name.format(y=y, m=m, d=d)
        return self.dir / f"{self.name}_{y}{m:02d}{d:02d}.xlsx"

    def parse_date(self, name: str):
        """name 传完整文件名（含扩展名），正则里请把扩展名也写上"""
        m = self.pattern.match(name)
        if not m:
            return None
        g = m.groupdict()
        try:
            now = datetime.now()
            y = int(g["y"]) if "y" in g and g.get("y") else now.year
            mo = int(g["m"])
            dd = int(g["d"])
            return datetime(y, mo, dd).date()
        except Exception:
            return None

    def exact(self, day: str) -> Path:
        return _resolve_dup(self.std_path(day))

    def list_all(self):
        out = []
        if not self.dir.exists():
            return out
        for f in self.dir.iterdir():
            if not f.is_file() or f.name.startswith(IGNORE_PREFIX):
                continue
            dt = self.parse_date(f.name)
            if dt:
                out.append((dt, f))
        return sorted(out, key=lambda t: t[0])

    def pick(self, day: str):
        """返回 (path, 是否回溯, 实际日期)。当日缺失且允许回溯时，取日期更晚的最近一份。"""
        day_d = datetime.fromisoformat(day).date()
        p = self.exact(day)
        if p.exists():
            return p, False, day_d
        if self.allow_newer:
            newer = [(d, f) for d, f in self.list_all() if d >= day_d]
            if newer:
                d, f = newer[0]
                return f, True, d
        return self.std_path(day), False, day_d


SOURCES = [Source(s) for s in CFG.get("sources", [])]


def check_ready(day: str):
    """所有源都取到即齐备。返回 (ok, missing 列表, ready 标记路径)"""
    rp = TRIGGER / f"ready_{day}.json"
    picks, missing, unstable = [], [], []
    for s in SOURCES:
        p, fb, cov = s.pick(day)
        if not p.exists():
            missing.append(f"{s.name}:{p.name}")
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
        if fb:
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


def _run_cli(prompt, out_log: Path, timeout, title, extra_args=None):
    """统一 CLI 调用：无窗口 + stdin 断开 + 强制精简产物。
    日志只回显权限参数，绝不回显提示词（否则关键字会污染自检判定）。"""
    import subprocess
    c = _cli_cfg()
    perm = list(extra_args or c["variants"][0])
    cmd = [c["node"], c["entry"], "-p", prompt] + perm + ["--output-format", "text"]
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    with open(out_log, "a", encoding="utf-8") as fh:
        fh.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} {title} =====\n")
        fh.write(f"[args] {' '.join(perm)}\n")
        fh.flush()
        return subprocess.run(cmd, cwd=str(ROOT), stdin=subprocess.DEVNULL,
                              stdout=fh, stderr=subprocess.STDOUT,
                              env=_cli_env(), creationflags=flags, timeout=timeout)


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
_RUN_SERIAL = threading.Lock()


def run_task(day: str):
    """跑一次任务。同一天跑过就跳过；成功的标记写 _trigger/_done/<day>.ok"""
    done = TRIGGER / "_done" / f"{day}.ok"
    done.parent.mkdir(parents=True, exist_ok=True)
    if done.exists():
        log(f"跳过：{day} 已完成")
        return
    out_log = LOGDIR / f"run-{day}.log"
    try:
        log(f"开始执行：{day}｜输出 {out_log.name}")
        r = _run_cli_retry(build_prompt(day), out_log, CFG["run_timeout"],
                           f"执行 {day}", extra_args=_perm_args())
        if r.returncode == 0:
            done.write_text(datetime.now().isoformat(), encoding="utf-8")
            log(f"执行完成：{day}（exit 0）")
            notify("任务完成", f"{day} 已执行，日志 {out_log.name}")
        else:
            log(f"执行失败：{day}（exit {r.returncode}）｜详见 {out_log.name}")
            notify("任务失败", f"{day} exit {r.returncode}")
    except Exception as e:
        log(f"执行异常：{day}｜{e}")


def run_task_serial(day: str):
    """按日串行，保证前一日先完成（后一日常依赖前一日的产物）"""
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

    warned = set()
    try:
        while True:
            try:
                today = datetime.now().date()
                days = {(today - timedelta(days=i)).isoformat()
                        for i in range(CFG.get("lookback_days", 1) + 1)}
                for day in sorted(days):
                    ok, missing, rp = check_ready(day)
                    if ok and day not in warned:
                        warned.add(day)
                        log(f"源文件齐备（{day}），标记 {rp.name}")
                        if CFG.get("auto_run"):
                            threading.Thread(target=run_task_serial, args=(day,), daemon=True).start()
                        else:
                            notify("源文件已齐备", f"{day} 可执行任务（auto_run=false，需人工触发）")
                    elif not ok:
                        warned.discard(day)
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
        day = sys.argv[2] if len(sys.argv) > 2 else datetime.now().date().isoformat()
        ok, missing, rp = check_ready(day)
        log(f"单次检查 {day}：{'齐备' if ok else '缺 ' + ','.join(missing)}")
        print(build_prompt(day))
        sys.exit(0)
    main()
