#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文本格式化工具 + 列宽配置持久化
"""

import json
from pathlib import Path

from rich.markup import escape as rich_escape

_CONFIG_DIR = Path.home() / ".exceltui"
_COL_WIDTHS_FILE = _CONFIG_DIR / "column_widths.json"


# ---------- 列宽持久化 ----------

def loadColWidthsConfig() -> dict:
    try:
        if _COL_WIDTHS_FILE.exists():
            return json.loads(_COL_WIDTHS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def saveColWidthsConfig(data: dict) -> None:
    try:
        _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        _COL_WIDTHS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


# ---------- 文本显示宽度 ----------

def displayWidth(s: str) -> int:
    """计算字符串在终端中的显示宽度（CJK=2，英文=1）"""
    w = 0
    for c in s:
        if (
            "\u4e00" <= c <= "\u9fff"
            or "\uff00" <= c <= "\uffef"
            or c in "，。！？、；：\u201c\u201d\u2018\u2019（）【】"
        ):
            w += 2
        else:
            w += 1
    return w


def padToDisplayWidth(s: str, width: int, truncate: bool = True) -> str:
    """按显示宽度左对齐填充或截断，保证终端对齐"""
    if truncate and displayWidth(s) > width:
        result = []
        cur = 0
        for c in s:
            if cur + displayWidth(c) > width - 2:
                result.append("..")
                break
            cur += displayWidth(c)
            result.append(c)
        s = "".join(result)
    return s + " " * (width - displayWidth(s))


def padToDisplayWidthRight(s: str, width: int) -> str:
    """按显示宽度右对齐填充"""
    return " " * (width - displayWidth(s)) + s


# ---------- 单元格值格式化 ----------

def formatDisplayValue(val) -> str:
    """格式化显示值：1.0 -> 1，并移除换行符防止 wrap"""
    if val is None:
        return ""
    s = str(val).strip().replace("\n", " ").replace("\r", " ")
    if not s:
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except ValueError:
        pass
    return s


def escapeForRich(s: str) -> str:
    """转义 Rich 标记字符（[、] 等），避免 MarkupError"""
    return rich_escape(s)
