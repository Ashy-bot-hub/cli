# -*- coding: utf-8 -*-
"""守护进程：watcher.py 崩掉后自动把它拉起来。

为什么需要这个
--------------
无人值守的看门狗最怕「静默失效」：pythonw 退出是无声的 —— 没有窗口、没有
弹窗，日志停在最后一行，源文件照常往目录里落，但没人处理。开机自启只在
登录时跑一次，所以崩了要等到下次重启才恢复，中间那段时间全是空白。

判活方式
--------
watcher.py 每轮轮询都会更新 `_trigger/_heartbeat`（节流 60 秒）。本脚本看它的
mtime，而不是只看 pid 在不在 —— pid 会被系统复用，心跳不会说谎，而且能顺带
发现「进程还在但已经卡死」这一类比崩溃更隐蔽的情况。

安全边界（刻意保守）
--------------------
* 只拉起，**不杀任何进程**。宁可少做不可误杀：它可能正在跑几小时的整周期重算。
* 心跳过期但 pid 仍存活 → 判定疑似卡死，只累计警告，连续多轮才升级措辞。
* 拉起失败按指数退避到 max_backoff，避免疯狂重启把机器拖死。
* 每轮重读配置，改参数不必重启守护；任何异常都只记日志、继续下一轮。

用法
----
把这个文件交给 pythonw.exe 运行（不要用 python.exe，否则会挂着黑窗口）：
    pythonw.exe watchdog.pyw
建议由 vbs 启动器或开机自启项调用，别直接双击 —— 双击在某些环境下会用
python.exe 而不是 pythonw.exe 关联。
"""
import ctypes
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
CONFIG = SCRIPTS / "config.json"
TARGET = SCRIPTS / "watcher.py"

_STILL_ACTIVE = 259
_QUERY_LIMITED = 0x1000
_NO_WINDOW = 0x08000000


def _dirs():
    """目录与配置一起读，避免调用方重复解析。"""
    cfg = {}
    try:
        cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        pass
    logdir = ROOT / cfg.get("log_dir", "_logs")
    trigger = ROOT / cfg.get("trigger_dir", "_trigger")
    return cfg, logdir, trigger


def log(msg: str):
    try:
        _, logdir, _ = _dirs()
        logdir.mkdir(parents=True, exist_ok=True)
        p = logdir / f"watchdog-{datetime.now():%Y-%m}.log"
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


def load_cfg() -> dict:
    cfg, _, _ = _dirs()
    c = cfg.get("watchdog") or {}
    return {
        "enabled": bool(c.get("enabled", True)),
        "interval": max(30.0, float(c.get("interval_seconds", 300))),
        "stale": max(120.0, float(c.get("stale_seconds", 900))),
        "max_backoff": max(60.0, float(c.get("max_backoff_seconds", 1800))),
        "stall_rounds": max(2, int(c.get("stall_warn_rounds", 3))),
    }


def pid_alive(pid: int) -> bool:
    """pid 是否指向仍在运行的进程。纯 ctypes，不依赖 psutil。"""
    if not pid or pid <= 0:
        return False
    try:
        h = ctypes.windll.kernel32.OpenProcess(_QUERY_LIMITED, False, int(pid))
        if not h:
            return False
        try:
            code = ctypes.c_ulong()
            if not ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code)):
                return False
            return code.value == _STILL_ACTIVE
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
    except Exception:
        return False


def _read_pid(p: Path) -> int:
    try:
        return int(p.read_text(encoding="utf-8").strip().splitlines()[0])
    except Exception:
        return 0


def _interp() -> Path:
    """优先用 pythonw（无黑窗口），退而用当前解释器。"""
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    return pyw if pyw.exists() else exe


def launch() -> bool:
    if not TARGET.exists():
        log(f"拉起失败：找不到 {TARGET}")
        return False
    try:
        subprocess.Popen(
            [str(_interp()), str(TARGET)],
            cwd=str(ROOT),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=_NO_WINDOW,
            close_fds=True,
        )
        return True
    except Exception as e:
        log(f"拉起异常：{e}")
        return False


def check_once(c: dict, state: dict, trigger=None) -> float:
    """检查一轮，返回下一轮应等待的秒数。state 里带 stall / fails 计数器。"""
    if trigger is None:
        _, _, trigger = _dirs()
    beat = trigger / "_heartbeat"
    w_lock = trigger / "_watcher.lock"

    pid = _read_pid(w_lock)
    alive = pid_alive(pid)
    try:
        age = time.time() - beat.stat().st_mtime
    except Exception:
        age = None
    fresh = age is not None and age < c["stale"]
    desc = "心跳文件不存在" if age is None else f"心跳已 {age / 60:.0f} 分钟未更新"

    if fresh:
        if state["stall"] or state["fails"]:
            log(f"心跳正常（pid {pid}，{age / 60:.1f} 分钟前更新），守护转入常态")
        state["stall"] = 0
        state["fails"] = 0
        return c["interval"]

    if alive:
        state["stall"] += 1
        if state["stall"] == 1:
            log(f"警告：pid {pid} 存活但{desc}，可能是长任务卡住，先观察不处理")
        elif state["stall"] == c["stall_rounds"]:
            log(f"严重：pid {pid} 连续 {state['stall']} 轮无心跳（{desc}），"
                f"疑似卡死，需人工确认；守护按约定不自动杀进程")
        return c["interval"]

    state["stall"] = 0
    state["fails"] += 1
    wait = min(c["max_backoff"], c["interval"] * (2 ** min(state["fails"] - 1, 4)))
    if launch():
        log(f"主进程不在（lock pid {pid or '无'} 已退出，{desc}），已拉起；"
            f"第 {state['fails']} 次，{wait:.0f}s 后复查")
    else:
        log(f"拉起失败，第 {state['fails']} 次，{wait:.0f}s 后重试")
    return wait


def main():
    cfg = load_cfg()
    if not cfg["enabled"]:
        log("守护未启用（watchdog.enabled=false），退出")
        return
    _, _, trigger = _dirs()
    wd_lock = trigger / "_watchdog.lock"

    # 守护自身单实例：有存活守护就不再开第二个
    old_wd = _read_pid(wd_lock)
    if old_wd and old_wd != os.getpid() and pid_alive(old_wd):
        log("已有守护进程在运行，本次退出（单实例保护）")
        return
    try:
        trigger.mkdir(parents=True, exist_ok=True)
        wd_lock.write_text(
            f"{os.getpid()}\n{datetime.now().isoformat(timespec='seconds')}\n",
            encoding="utf-8",
        )
    except Exception as e:
        log(f"写守护锁失败，退出：{e}")
        return

    log(f"守护启动｜pid {os.getpid()}｜间隔 {cfg['interval']:.0f}s"
        f"｜心跳超时 {cfg['stale']:.0f}s｜退避上限 {cfg['max_backoff']:.0f}s")
    state = {"stall": 0, "fails": 0}

    try:
        wait = check_once(cfg, state, trigger)
    except Exception as e:
        log(f"首轮检查异常：{e}")
        wait = cfg["interval"]

    while True:
        time.sleep(wait)
        try:
            cfg = load_cfg()
            if not cfg["enabled"]:
                log("守护已被关闭（watchdog.enabled=false），退出")
                return
            wait = check_once(cfg, state, trigger)
        except Exception as e:
            log(f"检查异常：{e}")
            wait = cfg["interval"]


if __name__ == "__main__":
    main()
