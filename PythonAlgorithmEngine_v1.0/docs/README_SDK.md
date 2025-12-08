# PythonAlgorithmEngine SDK — 使用手冊（v1.0.0）
## 1\. 簡介

PythonAlgorithmEngine SDK 是一套讓 C# (.NET 8) 直接呼叫 Python 實作之演算法 的整合框架。
專為自動化測試站與內部 PD（Photodiode）分析流程設計，支援：

* PD LIV 演算法
* PD Stability 演算法
* 圖片輸出（Matplotlib）
* LIV 與 Stability 的 PASS/FAIL 判定
* 跨 UI 架構使用（WPF / WinForms / Console）
SDK 將所有複雜的 Python 呼叫、資料建立、Plot、錯誤處理封裝起來，用戶端僅需：
1. 引入 DLL
2. 建立 config
3. 呼叫 AlgorithmEngine
即可取得計算結果與圖檔。

## 2\. SDK 內容物（目錄結構）

SDK 壓縮包（例如：`PythonAlgorithmEngine_SDK_v1.0.0.zip`）內容如下：
```text
PythonAlgorithmEngine_SDK_v1.0.0/
 ├─ lib/
 │   └─ PythonAlgorithmEngine.Core.dll
 │
 ├─ pythonAlgo/
 │   ├─ algorithm_dispatcher.py
 │   ├─ pd_analysis.py
 │   └─ （其他 .py 演算法模組）
 │
 ├─ config/
 │   └─ python_engine_config_template.json
 │
 ├─ samples/
 │   └─ WinFormsDemo/
 │       ├─ WinFormsDemo.sln
 │       ├─ WinFormsDemo.csproj
 │       ├─ Form1.cs
 │       ├─ Form1.Designer.cs
 │       ├─ Program.cs
 │       └─ python_engine_config.json（示範用）
 │
 └─ docs/
     └─ README_SDK.md
```

## 3\. 安裝需求
### 3.1 .NET

* 必須使用 .NET 8（WPF / WinForms / Console 均可）
* Windows 10 / 11

### 3.2 Python
* Python 3.8–3.12 皆可
* 需安裝下列套件：
```bash
pip install numpy matplotlib
```
如果你使用虛擬環境（venv），請確保設定檔中填寫的是 venv 內的 python.exe。

## 4\. 快速開始

以下是將 SDK 整合到 C# 專案的步驟。
### Step 1 — 複製 python/ 資料夾至你的電腦

例如：
```bash
D:\Algo\Python\
```

裡面會有：
* algorithm_dispatcher.py
* pd_analysis.py
* input/
* output/
注意：Python 腳本絕對不能遺漏或改錯路徑，C# Call Python 時會直接使用。

### Step 2 — 複製設定檔樣板到你的 C# 專案

從：
```bash
config/python_engine_config_template.json
```

複製到你自己的專案資料夾底下，改名：
```bash
python_engine_config.json
```
並在 File Properties 設定：
* Build Action：Content
* Copy to Output Directory：Copy if newer

### Step 3 — 修改設定檔中的路徑

`python_engine_config.json`：
```csharp
{
  "PythonExePath": "D:\\Algo\\Python\\venv\\Scripts\\python.exe",
  "DispatcherPath": "D:\\Algo\\Python\\algorithm_dispatcher.py",
  "InputDirectory": "D:\\Algo\\Python\\input"
}
```

三個欄位說明：

欄位					說明
`PythonExePath`		Python 可執行檔（可填 venv）
`DispatcherPath`	主控演算法的 python 腳本
`InputDirectory`	LIV / Stability 的輸入資料暫存路徑

### Step 4 — 將 DLL 加入 C# 專案

在你的 C# 專案：
```bash
右鍵 Dependencies → Add Reference → Browse...
```

選擇：
```bash
lib/PythonAlgorithmEngine.Core.dll
```
加入後，就可以在程式碼使用 SDK。

### Step 5 — 在程式碼加入 using
```csharp
using PythonAlgorithmEngine.Core.Config;
using PythonAlgorithmEngine.Core.Engine;
using PythonAlgorithmEngine.Core.AlgorithmInput;
```

### Step 6 — 初始化 AlgorithmEngine（SDK 入口）
```csharp
string configPath = Path.Combine(
    AppDomain.CurrentDomain.BaseDirectory,
    "python_engine_config.json");

var config = PythonEngineConfig.LoadFromFile(configPath);
var engine = new AlgorithmEngine(config);
```

## 5\. 演算法呼叫範例
### 5.1 呼叫 PD LIV
```csharp
var currents = new List<double> { 0, 0.5, 1, 1.5, ... };
var powers   = new List<double> { 0, 0,  1.1, 1.6, ... };
double thresholdPower = 1.1;

var rules = new PdLivThresholdRules
{
    SlopeLSL = 0.6,
    MaxPowerLSL = 11.0,
    ThresholdPowerLSL = 1.0,
    ThresholdPowerUSL = 1.5,
};

var result = await engine.RunPdLivAsync(
    currents,
    powers,
    thresholdPower,
    rules
);

result.ThrowIfError();   // Python 錯誤會直接丟出例外

Console.WriteLine($"PASS = {result.PassStatus}");
Console.WriteLine("Plot = " + result.PlotPath);
```

### 5.2 呼叫 PD Stability
```csharp
var samplePowerArray = new List<double> 
{
    12.80, 12.81, 12.79, ...
};

double usl = 1.0;

var result = await engine.RunPdStabilityAsync(
    samplePowerArray,
    usl
);

result.ThrowIfError();

Console.WriteLine($"PASS = {result.PassStatus}");
Console.WriteLine("Plot = " + result.PlotPath);
```

## 6\. PD LIV 與 Stability 回傳資料格式（PythonAlgorithmResult）

SDK 的結果物件：
```csharp
public class PythonAlgorithmResult
{
    public bool PassStatus { get; set; }
    public Dictionary<string, object>? CalculatedValues { get; set; }
    public string PlotPath { get; set; }
    public string Error { get; set; }
}
```

你可以讀到：
```csharp
if (result.CalculatedValues != null)
{
    foreach (var kv in result.CalculatedValues)
        Console.WriteLine($"{kv.Key} = {kv.Value}");
}
```

## 7\. Windows Forms 範例（顯示圖片）

圖片輸出會覆蓋：
```bash
python/output/pd/*.png
```

WinForms 載圖方式建議使用 temp file：
```csharp
using (var fs = new FileStream(result.PlotPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    pictureBox.Image = Image.FromStream(fs);
}
```

SDK sample 裡的 WinForms Demo 已示範完整流程。

## 8\. SDK 內部架構
```text
PythonAlgorithmEngine.Core/
 ├─ Config/
 │   └─ PythonEngineConfig.cs
 │
 ├─ AlgorithmInput/
 │   ├─ PdLivInputBuilder.cs
 │   └─ PdStabilityInputBuilder.cs
 │
 └─ Engine/
     ├─ AlgorithmEngine.cs        ← Facade（外部唯一入口）
     ├─ PdAlgorithmEngine.cs      ← internal
     ├─ PythonAlgorithmManager.cs ← internal（溝通 Python）
     └─ PythonAlgorithmResult.cs
```

外部使用者只會用到：
* AlgorithmEngine
* PythonEngineConfig
* PythonAlgorithmResult
* PdLivThresholdRules
* PdStabilityThresholdRules
其他類別已 internal，不需操作。

## 9\. 常見錯誤排查
錯誤											原因						解法
python_engine_config.json not found	JSON 	沒複製到output			設定 Copy to Output: Copy if newer
Cannot find python.exe						PythonExePath 錯		填寫正確完整路徑
No module named numpy						Python 套件未安裝		重新執行 pip install
Plot 無法顯示								圖片被鎖定				用 temp file 載入圖片
Python 無法執行								路徑含中文或空白			建議使用全英文路徑

## 10\. 版本管理
```text
PythonAlgorithmEngine_SDK_v1.0.0.zip
```
v1.0.0 — 第一次穩定版本


## 11\. 聯絡 / 維護者

Maintainer: Harris

Role: Firmware Engineer