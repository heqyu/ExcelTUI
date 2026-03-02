#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/CSV TUI 工具 - 终端界面只读查看 Excel/CSV 文件
支持: Sheet 选择、单元格浏览、行详情、搜索、过滤
"""

import argparse
import sys
from pathlib import Path

from exceltui.screens import ExcelTuiApp
from exceltui.schema import findSchemaConfig, loadSchemaConfig


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel/CSV TUI - 终端界面只读查看表格文件")
    parser.add_argument("file", type=str, help="文件路径 (.xlsx/.xlsm/.xls/.csv)")
    parser.add_argument(
        "--config", "-c", type=str, default=None,
        help="xls2lua schema .py 配置文件路径（不指定时自动发现）",
    )
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"Error: file not found: {path}", file=sys.stderr)
        return 1

    supported = {".xlsx", ".xlsm", ".xls", ".csv"}
    if path.suffix.lower() not in supported:
        print(
            f"Error: unsupported format '{path.suffix}', supported: {', '.join(sorted(supported))}",
            file=sys.stderr,
        )
        return 1

    schemaData: dict | None = None
    configPath = args.config or findSchemaConfig(str(path.resolve()))
    if configPath:
        schemaData = loadSchemaConfig(configPath) or None

    app = ExcelTuiApp(str(path.resolve()), schemaData=schemaData)
    result = app.run()
    return result if isinstance(result, int) else 0


if __name__ == "__main__":
    sys.exit(main())
