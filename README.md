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
你收到文件 ─┬─ 另存到 source_a\ / source_b\（几个源都行）
            └─ 留在下载目录 / 聊天软件的文件目录（由 external_sources 自动认领，复制过来）
                    │
           watcher.py（pythonw 常驻，默认 5 秒轮询）
                    │  ① 文件名归一（兼容 (1)（1）- 副本，匹配前先去掉后缀）
                    │  ② 写入稳定性检查（防读到复制中的半截文件）
                    │  ③ 同日多份取 mtime 最新（修正版覆盖旧版，旧版进 _logs\_history\）
                    │  ④ 补日窗口：最近 N 天逐日判，已出产物/已跑过的跳过
                    │  ⑤ 齐备判定（可按需回溯取更晚日期的累计表）
                    ▼
           _trigger\ready_{日期}.json   ← 证据留痕
                    │
                    ▼
           按日期升序入 FIFO 队列（单 worker 线程，绝不并发写同一产物）
                    │
                    ▼
           调 WorkBuddy 内置 CLI（-p 非交互｜--output-format stream-json 实时流）
                    │  权限：acceptEdits + -y + --allowedTools "Bash(*)"
                    │  环境：强制精简产物 + 指定解释器绝对路径
                    ▼
           按 task_spec.md 执行
                    ├─ _logs\run-{日期}.log        完整执行流水（实时刷新）
                    ├─ _logs\_stream\*.jsonl       原始事件流（排障）
                    └─ conclusions\{日期}.md       结论摘要（人直接看这个）
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

只跑外部来源取件（不判定、不执行）：

```bash
python scripts/watcher.py --pickup
```

它会列出各外部目录里认领到的候选文件与搬运结果，用来确认 `external_sources` 写对了。

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
├─ conclusions/                  # 每次任务的结论摘要（自动创建）
├─ task_spec.md                  # 给 agent 的业务说明书
├─ _logs/                        # 运行日志（自动创建）
│  ├─ watcher-YYYY-MM.log        # 主日志（到件、入队、等待心跳）
│  ├─ run-{日期}.log             # 单次执行流水（实时刷新）
│  ├─ preflight-{日期}.log       # 只读自检
│  ├─ startup.log                # 开机自启痕迹（仅 Startup 版 vbs 写）
│  ├─ _stream/*.jsonl            # CLI 原始事件流
│  └─ _history/                  # 被替换的旧版源文件
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
| `lookback_days` | 补日窗口：今天往前几天（默认 7）。晚到超过这个天数的源文件不再自动补 |
| `wait_log_interval_seconds` | 缺件等待日志的心跳间隔（默认 1800s）。整天没文件时靠它证明监听还活着 |
| `produced_glob` | **产物判定**：该日期已有匹配文件就永不重跑，支持 `{y}{m}{d}`（不补零）。放宽 `lookback_days` 时强烈建议配置 |
| `conclusion_dir` | 结论摘要目录（默认 `conclusions`），每次任务一文件 |
| `external_sources[]` | 外部来源取件（下载目录、聊天软件文件目录）：只复制不移动，按 mtime 取最新 |
| `auto_run` | `true` 自动执行；`false` 只识别 + 提示，由人触发 |
| `preflight_on_start` | 启动时做一次只读自检 |
| `run_timeout` | 单次任务超时秒数 |
| `task.spec_doc` / `task.prompt` | 业务说明书路径；提示词模板，支持 `{root}` `{spec}` `{day}` `{notes}` |
| `cli.*` | CLI 的 node、入口脚本、Git Bash、`output_format`（`stream-json` 推荐）、权限组合候选 |

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

### 文件永远不落到源目录怎么办（外部来源取件）

最常见的坑：你以为「文件到了」，其实它躺在浏览器下载目录或聊天软件的文件目录里，
**源目录压根没有这个文件**。只监听源目录的写法在这里会一直等，日志却看似正常。

`external_sources` 解决这个问题：按 `sources[].pattern` 去外部目录里认领，复制进源目录。

```json
"external_sources": [
  {
    "name": "微信文件",
    "enabled": true,
    "root": "D:/Backup/Documents/xwechat_files",
    "depth": 5,
    "max_age_days": 8,
    "interval_seconds": 30,
    "sources": ["生产日报", "出入库"]
  }
]
```

| 字段 | 说明 |
| --- | --- |
| `root` | 外部目录根 |
| `depth` | 从 root 往下最多走几层（root 的直接子项算第 1 层）。微信 4.x 是 `<root>\<wxid>\msg\file\<YYYY-MM>\`，需要 ≥ 4 |
| `max_age_days` | 只认领最近几天落盘的文件，建议 ≥ `lookback_days` |
| `interval_seconds` | 扫描间隔，外部目录文件多，别设太密 |
| `sources` | 限定用哪几个源的正则去匹配；留空表示全部 |

三条硬规则：
1. **只 `copy2`，绝不 move/delete** —— 外部目录归别的程序管，动了就出事。
2. **匹配前先去掉 `(1)`、`（1）`、`- 副本`** —— 否则第二次下载的文件名匹配不上，「取最新」永远看不到第二份。
3. **同一「源+日期」有多份时只取 mtime 最新的一份**。

排查用 `python scripts/watcher.py --pickup`，它会列出候选与搬运结果。

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

### 附：五条工程经验

- **上游模型会瞬时故障**（`500 temporarily unavailable` / `429`）。只对这类错误重试（默认 3 次），业务性失败不重试。
- **补日必须按日期升序串行**：错过几天后可能同时命中多天，两个 CLI 同写一个产物会互相覆盖。
  注意「抢锁」并不能保证顺序 —— 后一天产物依赖前一天产物，顺序错会算错累计数。
  本项目改为 **FIFO 队列 + 单 worker 线程**，入队前按日期排序。
- **`--output-format text` 是全缓冲**：重定向到文件时长任务中途一个字节都不写（实测跑 4 分钟日志始终 143 字节），
  只能靠进程存活 + CPU 增量猜它在不在干活。改用 `--output-format stream-json` 逐行读取、逐行 flush。
- **stream-json 里解析不了的行必须原样写出**：stderr、Node 警告、错误堆栈都在这些行里，
  瞬时故障重试的正则就是靠它们判定的。吞掉就等于把重试机制废了。
- **判定「谁是新版」不能只比文件大小**：对方发个修正版、只改一个数字，字节数常常完全一致。
  正确顺序是「大小不同→比 mtime；大小相同→再比 SHA256；内容不同且源更新→替换，旧版留档」。

### 7~10 条同源坑（同上，一并有代码修法）

7. **放宽补日窗口会推翻历史结果**：窗口一放大，那些「源早已齐备且产物早已生成」的旧日期会被当缺件重跑。
   > 修法：`produced_glob` 做产物判定，命中即跳过。`{m}` 不补零时注意前缀撞号（`10.1` vs `10.11`），
   > 让模板后面紧跟非数字字符（如 `日`）即可避免。
8. **陈旧执行锁会把某个日期永久堵死**：断电或进程被杀时 `_trigger/_running/<日期>.lock` 不会被清，
   之后每次都判「正在执行中」。
   > 修法：锁存在时间超过 `run_timeout + 300` 秒即判定残骸并清理。
9. **整天没文件时日志一片空白**：分不清是「守在等」还是「早就挂了」。
   > 修法：缺件首次出现记一行，之后按 `wait_log_interval_seconds` 打心跳，明确写「监听正常」。
10. **VBS 启动器的四个坑**（都踩过）：
    - `.vbs` 按 ANSI 读取，中文路径必须用 `ChrW()` 十进制编码；
    - **注释里也不能有中文**，交付前校验 `non_ascii == 0`；
    - `ChrW` 表达式要**反向解码验证**（解码后 `os.path.exists`），能抓出漏下划线、少一个字这类肉眼看不见的错；
    - 冷启动要留重试余量（本项目 12 次 × 10 秒），否则登录瞬间磁盘未就绪就写 E3 退出。

---

## 验活清单

- [ ] `_logs/watcher-*.log` 出现 `启动｜…｜pid xxxx`，与任务管理器里的进程号一致
- [ ] 有 `任务执行线程已启动｜补日窗口 N 天`（证明 FIFO 队列已起）
- [ ] `_trigger/_watcher.lock` 内容为该 pid
- [ ] `_logs/preflight-*.log` 通过（通道 + 权限/python 两段）
- [ ] `_trigger/_perm_mode.json` 已生成
- [ ] 有 `xxxx-xx-xx 等待源文件：缺 …` 一行（证明监听活着，之后每 30 分钟一条心跳）
- [ ] 往源目录丢当日文件 → `_trigger/ready_*.json` 生成，日志出现 `已入队`
- [ ] `_logs/run-*.log` **逐行刷新**（不是跑完才一次性写出）
- [ ] 产物已更新，且 `conclusions/{日期}.md` 已生成
- [ ] 重启机器后 `_logs/startup.log` 出现最新时间戳（仅装了开机自启版才检查）

---

## 已知限制

- 仅 Windows（用了 `CREATE_NO_WINDOW`、`ctypes` 取进程状态、`winsound`）。
- 依赖 WorkBuddy 桌面版内置 CLI；其路径与产物名可能随版本变化，升级后建议重跑一次自检。
- `auto_run=true` 时 CLI 无人值守执行，**能写文件、能跑命令**。不能接受就设 `auto_run=false`，退回「只提醒」。
- 本项目不包含任何公司/客户/价格数据，`task_spec.example.md` 只是格式示例。

---

## 更新记录

- 首版：就地监听、齐备判定、CLI 非交互链路（六坑）、只读自检、开机自启。
- 第二版：
  1. `external_sources` 外部来源取件（下载目录 / 聊天软件文件目录，只复制不移动）。
  2. 冲突判定改按 mtime 取最新 + SHA256 兜底，修正版不再被当重复丢掉；匹配前去掉 `(1)` 类后缀。
  3. 补日窗口 `lookback_days`（默认 7）+ `produced_glob` 产物判定，放宽窗口不会重算历史。
  4. FIFO 队列 + 单 worker，补多日严格按日期升序。
  5. `stream-json` 实时流 + `conclusions/{日期}.md` 结论摘要 + `_logs/_stream` 原始流存档。
  6. 缺件等待日志与心跳；陈旧执行锁自动清理。
  7. 开机自启 vbs 加固（解释器回退、冷启动重试、`startup.log` 痕迹）。

---

## 许可

MIT
