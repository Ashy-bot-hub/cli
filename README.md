# file-watcher-automation-template

**文件到齐 → 自动调用 WorkBuddy CLI 执行任务**（Windows / Python / 无窗口常驻）

一份可直接改的骨架：把「每天/每周固定要做的表格、报表、对账」变成**存完文件就不管**。
不含任何业务数据，只有通用的监听、判定、调用、重试、自检逻辑。

---

## 为什么不用聊天软件机器人

个人微信/企业微信的群机器人**不能接收消息**（WebHook 是单向推送），《微信个人账号使用规范》也禁止第三方工具接入客户端。
所以本项目**完全不接触聊天软件进程**，只监听「你另存文件」的那个目录 —— 合规且零封号风险。

---

## 架构

```
你收到文件 → 另存到 source_a\ / source_b\（几个源都行）
                    │
           watcher.py（pythonw 常驻，默认 5 秒轮询）
                    │  ① 文件名归一（兼容 (1)（1）- 副本）
                    │  ② 写入稳定性检查（防读到复制中的半截文件）
                    │  ③ 齐备判定（今日 + 可回溯的历史窗口）
                    ▼
           _trigger\ready_{日期}.json   ← 证据留痕
                    │
                    ▼
           调 WorkBuddy 内置 CLI（-p 非交互模式）
                    │  权限：acceptEdits + -y + --allowedTools "Bash(*)"
                    │  环境：强制精简产物 + 指定解释器绝对路径
                    ▼
           按 task_spec.md 执行 → 日志 + 提示
```

---

## 三步上手

1. **摆文件**（目录结构见下），把 `scripts/config.example.json` 复制为 `scripts/config.json`。
2. **改配置**：`sources` 里填你的目录和文件名正则；`cli` 里填本机 CLI 与 node 路径；
   `task.prompt` 改成你的提示词模板。
3. **启动**：双击 `scripts/start_watcher.vbs`（无窗口常驻）。
   首跑会在后台做一次只读自检，结果写 `_logs/preflight-{日期}.log`，出现 `PREFLIGHT_DONE` 或
   `预检完成` 即通过。

验证单日状态（不启动常驻）：

```bash
python scripts/watcher.py --once 2026-10-21
```

它会打印「齐备 / 缺什么」以及将要发给 CLI 的完整提示词，**不执行任务**，适合调试配置。

---

## 目录结构

```
<project>/
├─ scripts/
│  ├─ watcher.py                 # 主体
│  ├─ config.json                # 你的配置（由 config.example.json 复制）
│  ├─ start_watcher.vbs          # 无窗口启动
│  └─ startup_watcher.vbs        # 开机自启版（放「启动」文件夹）
├─ source_a/                     # 源 1
├─ source_b/                     # 源 2
├─ output/                       # 任务产物
├─ task_spec.md                  # 给 agent 的业务说明书
├─ _logs/                        # 运行日志（自动创建）
└─ _trigger/                     # ready 标记、锁、权限记录（自动创建）
```

---

## 配置字段

| 字段 | 说明 |
| --- | --- |
| `sources[].name` | 源名称，出现在日志与提示词里 |
| `sources[].dir` | 相对项目根的目录 |
| `sources[].pattern` | **对完整文件名**（含扩展名）匹配的正则，用命名分组 `y`/`m`/`d` 取日期；缺 `y` 按当年 |
| `sources[].std_name` | 标准文件名模板，如 `A_{m}.{d}.xlsx`；用于把 `A_10.21(1).xlsx` 归一 |
| `sources[].allow_newer_fallback` | 当日文件缺失时，允许用**日期更晚**的最近一份回溯取数（适合「周期累计堆叠」类报表） |
| `poll_seconds` / `stable_seconds` | 轮询间隔；文件需静置多少秒才算落盘完成 |
| `lookback_days` | 检查窗口：今天往前几天（补日场景设 1~3） |
| `auto_run` | `true` 自动执行；`false` 只识别 + 提示，由人触发 |
| `preflight_on_start` | 启动时做一次只读自检 |
| `run_timeout` | 单次任务超时秒数 |
| `task.spec_doc` / `task.prompt` | 业务说明书路径；提示词模板，支持 `{root}` `{spec}` `{day}` `{notes}` |
| `cli.*` | CLI 的 node、入口脚本、Git Bash、权限组合候选 |

### 文件名不含日期怎么办

默认 `date_source: "filename"`，要求正则里用命名分组 `y`/`m`/`d` 取出日期（`y` 可省略，按当年）。

有些报表文件名固定、当天覆盖（例如 `银行流水.xlsx`），这种情况把该源设成：

```json
{
  "name": "银行流水",
  "dir": "source_c",
  "pattern": "^银行流水",        // 只用来判断「是不是我这类型文件」
  "date_source": "mtime"        // 日期取文件最后修改时间
}
```

也可以放在顶层当全局默认：`"date_source": "mtime"`，单个源写了就以源的为准。

注意：`mtime` 模式判断的是「这个文件最后一次被修改是哪天」，所以

- 文件被复制到目录时若保留了原始时间戳，日期会算成原始那一天（`copy` 而非 `move` 时会遇到）；
- 同一天多次覆盖同名文件，只会认最新那一份，不留历史版本。

如果这两条不可接受，说明你的场景不是「按天」，而是「按批次/按版本」，需要把 day 换成批次键（见下方「适配其他任务」）。

---

## WorkBuddy CLI 无人值守的六个坑（本项目的核心内容）

这些都是实际调试出来的，不是文档抄来的。抄代码时**不要改回裸调用**。

### 1. `.bat` 里的中文会乱码

`.bat` 以 UTF-8 保存中文，cmd 按系统 GBK 解析 → 路径失效、`exit 9009`（命令不存在）。

> 修法：`.bat` 内容保持**纯 ASCII**，路径用 `%~dp0` 运行时推导；提示词用英文或改走文件。

### 2. `Cannot find module '../dist/codebuddy'`

桌面版 CLI 的 `dist` 目录可能只带精简产物 `codebuddy-lite-wb.mjs`；
`-p` 模式先找 `codebuddy-headless` 再回退 `codebuddy`，两个都缺就报这个错。

> 修法：设环境变量 `CODEBUDDY_FORCE_LITE_WB_BUNDLE=1`（可再加 `CODEBUDDY_FORCE_HEADLESS_BUNDLE=1`）。

### 3. `exit 127: python: command not found`

CLI 的 bash 环境 **PATH 与你自己的终端不同**（你的终端有 python，它可能没有）。

> 修法：执行 Python 一律用解释器绝对路径，并把这条写进 `task_spec.md`。
> 判断依据只能来自 CLI 自己的输出，**不要用你的环境推断它的环境**。

### 4. `pythonw` 拉起 CLI 会弹黑窗并永久挂死

`pythonw` 没有控制台，子进程继承的 stdin 不可用；CLI 读 stdin 时永久阻塞，
同时 node.exe 作为控制台程序会新分配一个窗口。

> 修法：
> ```python
> subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=fh, stderr=subprocess.STDOUT,
>                creationflags=0x08000000)   # CREATE_NO_WINDOW
> ```

### 5. 权限模式 ≠ 放行 Bash

`--permission-mode acceptEdits` 只放行文件编辑；`-y`（dangerously-skip-permissions）对高风险操作仍会询问，
非交互环境下「询问」等于「拒绝」，于是模型回你一句 `SHELL_DENIED`。

> 修法：显式列出允许的工具，注意 `Bash` 要写通配形式：
> ```
> --permission-mode acceptEdits -y --allowedTools "Bash(*),Read,Write,Edit,Glob,Grep"
> ```
> 本项目会在启动自检里按 `perm_variants` 依次探测，把能跑通的那组写入 `_trigger/_perm_mode.json`。

### 6. 自检假阳性

如果日志里回显了完整命令行（含提示词），而判定又用裸关键字（如判断是否包含 `PY_OK`），
那么提示词里的字样会让自检**永远通过**。

> 修法：日志只回显参数、不回显提示词；判定用带版本号的严格正则 `^\s*PY_OK\s+\d`。

### 附：两条工程经验

- **上游模型会瞬时故障**（`500 temporarily unavailable` / `429`）。只对这类错误重试（默认 3 次），业务性失败不重试。
- **补日会并发**：错过几天后可能同时触发「补前一日 + 出当日」，两个 CLI 同写一个产物会互相覆盖 → 必须加串行闸门（本项目 `_RUN_SERIAL`），按日期升序排队。

---

## 验活清单

- [ ] `_logs/watcher-*.log` 出现 `启动｜…｜pid xxxx`，与任务管理器里的进程号一致
- [ ] `_trigger/_watcher.lock` 内容为该 pid
- [ ] `_logs/preflight-*.log` 通过（通道 + 权限/python 两段）
- [ ] `_trigger/_perm_mode.json` 已生成
- [ ] 往源目录丢当日文件 → `_trigger/ready_*.json` 生成
- [ ] `_logs/run-*.log` 有完整输出，产物已更新

---

## 已知限制

- 仅 Windows（用了 `CREATE_NO_WINDOW`、`ctypes` 取进程状态、`winsound`）。
- 依赖 WorkBuddy 桌面版内置 CLI；其路径与产物名可能随版本变化，升级后建议重跑一次自检。
- `auto_run=true` 时 CLI 无人值守执行，**能写文件、能跑命令**。不能接受就设 `auto_run=false`，退回「只提醒」。
- 本项目不包含任何公司/客户/价格数据，`task_spec.example.md` 只是格式示例。

---

## 许可

MIT
