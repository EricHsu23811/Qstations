#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
algorithm_dispatcher.py

統一 C# ↔ Python 的呼叫介面（支援 alias + Hierarchical 結構）。
"""

import argparse
import json
import importlib
import sys
import traceback
import os
from typing import Any, Dict, Optional

# ---------------------------------------------------------------------------
# 專案根目錄設定：確保能 import 到 pd.*, tx.*, rx.* 這些 package
# ---------------------------------------------------------------------------
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# ---------------------------------------------------------------------------
# Script alias 設定：對 C# / 外部提供簡單名稱
# ---------------------------------------------------------------------------
SCRIPT_ALIAS_MAP: Dict[str, str] = {
    "PD": "pd.pd_analysis",
    "TX": "tx.tx_analysis",
    "RX": "rx.rx_analysis",
}


def print_json_and_exit(payload: Dict[str, Any], exit_code: int = 0) -> None:
    """
    將結果以單一 JSON 行輸出到 stdout，給 C# 反序列化。
    這個函式最後會呼叫 sys.exit()，不會正常 return。
    """
    try:
        json_str = json.dumps(payload, ensure_ascii=False, default=str)
    except Exception as e:
        json_str = json.dumps(
            {
                "pass_status": False,
                "error": f"Failed to serialize result: {e}",
            },
            ensure_ascii=False,
        )

    print(json_str)
    sys.exit(exit_code)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="M5 Algorithm Dispatcher - unified entry point for C# GUI"
    )
    parser.add_argument(
        "--script",
        required=True,
        help="Script alias or module name, e.g. 'PD', 'TX', 'RX' or 'pd.pd_analysis'",
    )
    parser.add_argument(
        "--action",
        required=True,
        help="Sub-test action name, e.g. 'LIV', 'STABILITY', 'TX_FOO', ...",
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Input JSON file path",
    )
    args = parser.parse_args()

    # 1) 解析 script alias → module name
    script_key = (args.script or "").strip()
    module_name = SCRIPT_ALIAS_MAP.get(script_key, script_key)

    # 2) 動態 import 指定的演算法模組
    module: Optional[Any] = None
    try:
        module = importlib.import_module(module_name)
    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        error_payload = {
            "pass_status": False,
            "error": f"Cannot import script '{module_name}' (from '{script_key}'): {e}",
        }
        print_json_and_exit(error_payload, exit_code=1)

    # 到這裡 module 一定不是 None
    assert module is not None

    # 3) 檢查該模組是否有 run_test(action, input_path)
    run_test = getattr(module, "run_test", None)
    if not callable(run_test):
        error_payload = {
            "pass_status": False,
            "error": (
                f"Module '{module_name}' missing required function "
                "run_test(action: str, input_json_path: str)"
            ),
        }
        print_json_and_exit(error_payload, exit_code=1)

    # 4) 呼叫 run_test()
    result: Optional[Dict[str, Any]] = None
    try:
        # type: ignore[call-arg]
        result = run_test(args.action, args.input)
    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        error_payload = {
            "pass_status": False,
            "error": f"Exception while executing {module_name}.run_test(): {e}",
        }
        print_json_and_exit(error_payload, exit_code=1)

    # 到這裡 result 一定不是 None
    assert result is not None

    # 5) 輸出結果 JSON（pass_status 由 JSON 自己決定）
    print_json_and_exit(result, exit_code=0)


if __name__ == "__main__":
    main()
