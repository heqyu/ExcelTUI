#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xls2lua schema 配置解析 + 自动发现

配置文件格式（Python 文件）：
    cfg = (
        {
            "SheetInfo": ["文件名.xlsx", "Sheet名", "导出Lua表名", "主键字段"],
            "NameMap": [
                ["中文列名", "英文字段名", "类型"],
                ...
            ]
        },
        ...
    )

解析结果结构：
    {
        sheet名: {
            "exportName": str,
            "keyField":   str,
            "cnToEn":     {中文列名: (英文字段名, 类型), ...},
        },
        ...
    }
"""

from pathlib import Path


def loadSchemaConfig(path: str) -> dict:
    """解析 xls2lua .py 配置文件，返回按 sheet 名索引的 schema 字典"""
    try:
        with open(path, encoding="utf-8") as f:
            content = f.read()
        ns: dict = {}
        exec(content, {"__builtins__": {}}, ns)  # noqa: S102
        cfg = ns.get("cfg", ())
        result: dict = {}
        for entry in cfg:
            sheet_info = entry.get("SheetInfo", [])
            name_map = entry.get("NameMap", [])
            if len(sheet_info) >= 3:
                sheet_name = sheet_info[1]
                export_name = sheet_info[2]
                key_field = sheet_info[3] if len(sheet_info) > 3 else ""
                result[sheet_name] = {
                    "exportName": export_name,
                    "keyField": key_field,
                    "cnToEn": {
                        row[0]: (row[1], row[2])
                        for row in name_map
                        if len(row) >= 3
                    },
                }
        return result
    except Exception:
        return {}


def findSchemaConfig(xlsxPath: str) -> str | None:
    """按约定目录自动寻找 xlsx 对应的 xls2lua schema .py 配置文件"""
    p = Path(xlsxPath)
    stem = p.stem
    candidates = [
        p.parent / f"{stem}.py",
        p.parent / "xls2lua" / "config" / f"{stem}.py",
        p.parent.parent / "xls2lua" / "config" / f"{stem}.py",
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return None
