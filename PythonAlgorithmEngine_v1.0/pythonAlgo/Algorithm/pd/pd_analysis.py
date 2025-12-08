# pd/pd_analysis.py

import os
import json
import math
from typing import List, Dict, Any, Optional

import numpy as np
import matplotlib.figure


# ==============================================================================
# 1. Plotting Helper
# ==============================================================================

def create_and_save_plot(
    x_data: List[float],
    y_data: List[float],
    output_filename: str,
    title: str,
    xlabel: str,
    ylabel: str,
    x_fit: Optional[np.ndarray] = None,
    y_fit: Optional[np.ndarray] = None,
    fit_slope: Optional[float] = None,
    y_max: Optional[float] = None,
    annotation_text: Optional[str] = None,
    hline: Optional[float] = None,
    hline_label: Optional[str] = None,
    std_dev: Optional[float] = None,
) -> str:
    """
    建立並儲存圖檔，回傳圖檔的絕對路徑。
    注意：這裡使用 matplotlib.figure.Figure，而非 pyplot 的全域狀態。
    """
    try:
        fig = matplotlib.figure.Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)

        # 1) 實際量測點
        ax.plot(
            x_data,
            y_data,
            marker='o',
            linestyle='None',
            markersize=4,
            label='Actual Data'
        )

        # 2) LIV fitted curve
        if x_fit is not None and y_fit is not None:
            ax.plot(
                x_fit,
                y_fit,
                color='red',
                linestyle='--',
                linewidth=1,
                label='Fitted Curve'
            )

            if fit_slope is not None:
                ax.text(
                    0.05,
                    0.95,
                    f"Slope: {fit_slope:.3f} mW/mA",
                    transform=ax.transAxes,
                    fontsize=10,
                    verticalalignment='top',
                    bbox=dict(
                        facecolor='white',
                        alpha=0.8,
                        edgecolor='red',
                        boxstyle='round,pad=0.3',
                    )
                )

        # 3) Stability 平均線
        if hline is not None:
            try:
                ax.axhline(hline, color='green', linestyle=':', linewidth=0.8)

                if hline_label:
                    ax.text(
                        0.0,
                        hline,
                        hline_label,
                        transform=ax.get_yaxis_transform(),
                        fontsize=9,
                        color='white',
                        fontweight='bold',
                        verticalalignment='center',
                        horizontalalignment='right',
                        bbox=dict(
                            boxstyle="square,pad=0.3",
                            fc="green",
                            ec="none",
                            alpha=0.8,
                        ),
                    )
            except Exception as e:
                # 這裡避免 plot 洩漏錯誤到 stdout，僅簡單印出
                print(f"Plot warning (hline): {e}")

        # 4) 資訊框
        if annotation_text:
            ax.text(
                0.98,
                0.25,
                annotation_text,
                transform=ax.transAxes,
                fontsize=9,
                verticalalignment='top',
                horizontalalignment='right',
                bbox=dict(
                    facecolor="#ffffff",
                    alpha=0.9,
                    edgecolor='red',
                    boxstyle='round,pad=0.3',
                ),
            )

        # 基本外觀設定
        ax.set_title(title, fontsize=12)
        ax.set_xlabel(xlabel, fontsize=10)
        ax.set_ylabel(ylabel, fontsize=10)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(loc='lower right', fontsize=8)

        if y_max is not None:
            ax.set_ylim(bottom=0, top=float(y_max))
        else:
            ax.set_ylim(bottom=0)

        fig.tight_layout()

        # 找到 Algorithm 目錄
        base_dir = os.path.dirname(os.path.abspath(__file__))  # .../pd/
        root_dir = os.path.dirname(base_dir)  # .../Algorithm/

        # 新增 output/PD 目錄
        output_dir = os.path.join(root_dir, "output", "pd")
        os.makedirs(output_dir, exist_ok=True)


        output_path = os.path.join(output_dir, output_filename)
        # 儲存圖檔
        fig.savefig(output_path, format='png')
        abs_path = os.path.abspath(output_path)
        return abs_path

    except Exception as e:
        print(f"Error creating plot: {e}")
        return ""


# ==============================================================================
# 2. 核心演算法函式：LIV / STABILITY
# ==============================================================================

def analyze_liv(raw_data: Dict[str, Any], threshold_rules: Dict[str, float]) -> Dict[str, Any]:
    """
    PD 子測項：LIV Curve 分析
    """
    print("--- PD / LIV Analysis ---")

    currents = raw_data.get('currents', [])
    powers = raw_data.get('powers', [])
    threshold_power = raw_data.get('threshold_power', 0.0)
    rules = threshold_rules

    currents_np = np.array(currents, dtype=float)
    powers_np = np.array(powers, dtype=float)


    # derek_fix_v1.1 >>>
    # --- 修正: 只選取功率超過 threshold 的點來計算斜率 排除啟動時的非線性區域，使斜率計算更準確 ---
    fit_threshold_lsl = rules.get('threshold_power_LSL', 0)
    mask = powers_np > fit_threshold_lsl
    # <<< derek_fix_v1.1

    if np.sum(mask) > 2:
        slope, intercept = np.polyfit(currents_np[mask], powers_np[mask], 1)
        fit_line_y = slope * currents_np + intercept
        fit_line_y[fit_line_y < 0] = 0
    else:
        slope = 0.0
        intercept = 0.0
        fit_line_y = np.zeros_like(powers_np)

    try:
        max_power = float(np.max(powers_np)) if len(powers_np) > 0 else 0.0
        calculated_values = {
            "slope_mw_per_ma": round(slope, 3),
            "max_power_mw": round(max_power, 3),
            "threshold_power_mw": round(float(threshold_power), 3),
        }
    except Exception as e:
        return {"pass_status": False, "error": f"LIV calculation error: {e}"}

    # 畫圖
    filename = "liv_result.png"
    plot_path = create_and_save_plot(
        x_data=list(currents_np),
        y_data=list(powers_np),
        output_filename=filename,
        title="L-I Curve Analysis",
        xlabel="Current (mA)",
        ylabel="Power (mW)",
        x_fit=currents_np,
        y_fit=fit_line_y,
        fit_slope=slope,
    )

    # 規格判定
    try:
        pass_slope = calculated_values['slope_mw_per_ma'] >= rules['slope_LSL']
        pass_max = calculated_values['max_power_mw'] >= rules['max_power_LSL']
        pass_thresh = (
            rules['threshold_power_LSL']
            <= calculated_values['threshold_power_mw']
            <= rules['threshold_power_USL']
        )
        is_pass = all([pass_slope, pass_max, pass_thresh])
    except Exception as e:
        return {"pass_status": False, "error": f"LIV rule error: {e}"}

    return {
        "pass_status": is_pass,
        "calculated_values": calculated_values,
        "plot_path": plot_path,
    }


def analyze_stability(raw_data: Dict[str, Any], threshold_rules: Dict[str, float]) -> Dict[str, Any]:
    """
    PD 子測項：功率穩定度分析
    """
    print("--- PD / Stability Analysis ---")

    power_array = raw_data.get('power_array', [])
    rules = threshold_rules

    try:
        power_array_np = np.array(power_array, dtype=float)
        avg_power = float(np.mean(power_array_np)) if len(power_array_np) > 0 else 0.0
        std_dev = float(np.std(power_array_np)) if len(power_array_np) > 0 else 0.0

        if avg_power == 0:
            stability_percent = 0.0
        else:
            stability_percent = (std_dev / avg_power) * 100.0

        calculated_values = {
            "stability_percent": round(stability_percent, 3),
            "avg_power_mw": round(avg_power, 3),
            "std_dev_mw": round(std_dev, 3),
        }
    except Exception as e:
        return {"pass_status": False, "error": f"Stability calculation error: {e}"}

    filename = "stability_result.png"
    time_axis = list(range(len(power_array_np)))
    actual_max = float(np.max(power_array_np)) if len(power_array_np) > 0 else 0.0
    y_top = math.ceil(actual_max * 1.02) if actual_max > 0 else None

    info_text = (
        f"Avg Pwr: {calculated_values['avg_power_mw']:.3f} mW\n"
        f"STD Dev: {calculated_values['std_dev_mw']:.4f} mW\n"
        f"Stability: {calculated_values['stability_percent']:.3f}%"
    )

    plot_path = create_and_save_plot(
        x_data=time_axis,
        y_data=power_array,
        output_filename=filename,
        title="Power Stability Test",
        xlabel="Time (s)",
        ylabel="Power (mW)",
        y_max=y_top,
        annotation_text=info_text,
        hline=calculated_values['avg_power_mw'],
        hline_label=f"Avg: {calculated_values['avg_power_mw']:.2f}",
        std_dev=calculated_values['std_dev_mw'],
    )

    try:
        is_pass = calculated_values['stability_percent'] < rules['stability_USL']
    except Exception as e:
        return {"pass_status": False, "error": f"Stability rule error: {e}"}

    return {
        "pass_status": is_pass,
        "calculated_values": calculated_values,
        "plot_path": plot_path,
    }


# ==============================================================================
# 3. 子測項 wrapper：對應不同 action
# ==============================================================================

def _run_pd_liv_test(input_path: str) -> Dict[str, Any]:
    if not os.path.exists(input_path):
        return {
            "pass_status": False,
            "error": f"Input JSON file not found: {input_path}",
        }

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception as e:
        return {
            "pass_status": False,
            "error": f"Failed to read or parse input JSON '{input_path}': {e}",
        }

    if "raw_data" not in payload or "threshold_rules" not in payload:
        return {
            "pass_status": False,
            "error": f"{input_path} missing 'raw_data' or 'threshold_rules'.",
        }

    return analyze_liv(payload["raw_data"], payload["threshold_rules"])


def _run_pd_stability_test(input_path: str) -> Dict[str, Any]:
    if not os.path.exists(input_path):
        return {
            "pass_status": False,
            "error": f"Input JSON file not found: {input_path}",
        }

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception as e:
        return {
            "pass_status": False,
            "error": f"Failed to read or parse input JSON '{input_path}': {e}",
        }

    if "raw_data" not in payload or "threshold_rules" not in payload:
        return {
            "pass_status": False,
            "error": f"{input_path} missing 'raw_data' or 'threshold_rules'.",
        }

    return analyze_stability(payload["raw_data"], payload["threshold_rules"])


# ==============================================================================
# 4. 統一給 dispatcher / C# 用的入口
# ==============================================================================

def run_test(action: str, input_json_path: str) -> Dict[str, Any]:
    """
    統一給 algorithm_dispatcher.py / C# 呼叫的入口。

    - action: 由 C# 明確指定，例如 "LIV", "STABILITY"
    - input_json_path: 對應此子測項的 input JSON 路徑
    """
    act = (action or "").strip().upper()

    if act == "LIV":
        return _run_pd_liv_test(input_json_path)
    elif act == "STABILITY":
        return _run_pd_stability_test(input_json_path)
    else:
        return {
            "pass_status": False,
            "error": f"Unknown PD action: '{action}'",
        }


# ==============================================================================
# 5. 本機測試入口（給演算法工程師在 terminal 直接跑）
# ==============================================================================

if __name__ == "__main__":
    from pprint import pprint

    liv_file = "liv_input.json"
    stab_file = "pd_stability_input.json"

    print("=== Local Test: PD / LIV ===")
    if os.path.exists(liv_file):
        pprint(run_test("LIV", liv_file))
    else:
        print(f"  Warning: {liv_file} not found.")

    print("\n=== Local Test: PD / STABILITY ===")
    if os.path.exists(stab_file):
        pprint(run_test("STABILITY", stab_file))
    else:
        print(f"  Warning: {stab_file} not found.")
