# M5 Algorithm Engine - Interface Specification

本文件定義 **C\# GUI** 與 **Python  (`pd_analysis.py`)** 之間的通訊介面與整合需求。

## 1\. 系統需求 (Requirements)

### 運行環境

  * **Python Runtime**: 3.8 或以上版本。
  * **Python Libraries**: 需安裝以下套件：
    ```bash
    pip install numpy matplotlib
    ```

### 整合方式

  * **執行方式**: C\# 透過 Command Line 呼叫 Python 腳本：
    ```cmd
    python pd_analysis.py
    ```
  * **工作目錄**: C\# 必須確保 `pd_analysis.py` 與輸入的 `.json` 檔案位於**同一目錄**下執行。

-----

## 2\. 輸入介面 (Input Specifications)

C\# 需在呼叫 Python 前，產生以下 JSON 檔案。

### 2.1 LIV 測試 (L-I Curve)

  * **檔案名稱**: `liv_input.json`
  * **格式定義**:
    ```json
    {
      "raw_data": {
        "currents": [0.0, 0.5, 1.0, ...],  // 單位: mA (Array of double)
        "powers":   [0.0, 0.0, 1.1, ...],  // 單位: mW (Array of double)
        "threshold_power": 1.1             // 單位: mW (Double, 於 I-th 測得)
      },
      "threshold_rules": {                 // 客戶定義的規格
        "slope_LSL": 0.6,
        "max_power_LSL": 11.0,
        "threshold_power_LSL": 1.0,
        "threshold_power_USL": 1.5
      }
    }
    ```

### 2.2 穩定度測試 (Stability)

  * **檔案名稱**: `pd_stability_input.json`
  * **格式定義**:
    ```json
    {
      "raw_data": {
        "power_array": [12.80, 12.81, ...] // 單位: mW (Array of double, 共 30 筆)
      },
      "threshold_rules": {
        "stability_USL": 1.0               // 單位: % (Double)
      }
    }
    ```

-----

## 3\. 輸出介面 (Output Specifications)

Python 執行完畢後，會產出以下結果供 C\# 讀取。

### 3.1 圖檔輸出 (Image Files)

C\# 需讀取並顯示於 GUI 的 `PictureBox`。

| 測試項目 | 輸出檔名 | 說明 |
| :--- | :--- | :--- |
| **LIV Curve** | `liv_result.png` | 包含原始數據點 (藍色) 與擬合曲線 (紅色虛線)。 |
| **Stability** | `stability_result.png` | 包含趨勢圖、平均值標示 (Y軸綠色標籤) 與數據統計框。 |

### 3.2 數據輸出 (Console Output)

Python 會將計算結果以 **JSON 字串** 形式列印在 **Standard Output**。C\# 需捕捉 Console 輸出來解析結果。

**LIV 測試回傳範例:**

```text
... (Log messages) ...
[LIV Result from JSON]
{
  "pass_status": true,      // 最終判定結果 (True/False)
  "plot_path": "C:\\...\\liv_result.png", // 圖檔絕對路徑
  "calculated_values": {
    "slope_mw_per_ma": 1.096,
    "max_power_mw": 12.8,
    "threshold_power_mw": 1.1
  }
}
```

**穩定度測試回傳範例:**

```text
... (Log messages) ...
[Stability Result from JSON]
{
  "pass_status": true,
  "plot_path": "C:\\...\\stability_result.png",
  "calculated_values": {
    "avg_power_mw": 12.8,
    "std_dev_mw": 0.014,
    "stability_percent": 0.108
  }
}
```

-----

## 4\. 例外處理 (Error Handling)

若發生錯誤 (如檔案遺失、數據不足無法擬合)，Python 將回傳含有 `error` 欄位的 JSON：

```json
{
  "pass_status": false,
  "error": "具體的錯誤訊息描述 (例如: Data array is empty)"
}
```