# -*- coding: utf-8 -*-
"""把含中文的路径转成 VBScript 的 ChrW() 表达式，供启动文件夹里的 .vbs 使用。

背景：放在「启动」文件夹里的 .vbs 无法用 WScript.ScriptFullName 推导项目路径，
而 .vbs 又必须保持纯 ASCII（WScript 按 ANSI 读取，中文会乱码）。
解决办法：把路径拆成 ASCII 字面量 + ChrW(十进制码点) 拼接，源码里就没有中文了。

用法：
    python gen_chrw_path.py "D:\\我的项目\\scripts\\watcher.py"
输出：
    "D:\\" & ChrW(25105) & ChrW(30340) & ... & "\\scripts\\watcher.py"
把它整行粘到启动器里 `target = ` 后面即可。
"""

import sys


def to_chrw(expr: str) -> str:
    parts, buf = [], ""
    for ch in expr:
        if ord(ch) < 128:
            buf += ch
        else:
            if buf:
                parts.append('"' + buf.replace('"', '""') + '"')
                buf = ""
            parts.append(f"ChrW({ord(ch)})")
    if buf:
        parts.append('"' + buf.replace('"', '""') + '"')
    return " & ".join(parts)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    print(to_chrw(sys.argv[1]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
