using System;
using System.Diagnostics;

namespace M5.Core
{
    /// <summary>
    /// 負責啟動 / 停止 Python server 的 Process 管理。
    /// 可以重複 Start/Stop，不會跳出 CMD 視窗。
    /// </summary>
    public class PythonServerManager : IDisposable
    {
        private Process? _process;

        /// <summary>
        /// 目前 Python server 是否仍在執行中
        /// </summary>
        public bool IsRunning => _process != null && !_process.HasExited;

        /// <summary>
        /// 當 Python 有輸出 stdout 時觸發
        /// </summary>
        public event Action<string>? StdOutReceived;

        /// <summary>
        /// 當 Python 有輸出 stderr 時觸發
        /// </summary>
        public event Action<string>? StdErrReceived;

        /// <summary>
        /// 當 Python process 結束時觸發（傳出 ExitCode）
        /// </summary>
        public event Action<int>? Exited;

        /// <summary>
        /// 啟動 Python server。
        /// </summary>
        /// <param name="pythonExe">python.exe 的完整路徑</param>
        /// <param name="scriptPath">要執行的 .py 檔路徑</param>
        /// <param name="workingDirectory">工作目錄（通常是 script 所在資料夾）</param>
        /// <param name="arguments">額外參數（不含 scriptPath，本方法會自動加上）</param>
        /// <param name="redirectOutput">是否要重導 stdout/stderr 以便顯示在 GUI log</param>
        public void Start(
            string pythonExe,
            string scriptPath,
            string workingDirectory,
            string arguments,
            bool redirectOutput = false)
        {
            // 如果已經在跑，就不重複啟動
            if (IsRunning)
                return;

            var psi = new ProcessStartInfo
            {
                FileName = pythonExe,
                // 這裡把 scriptPath 與額外參數接起來
                Arguments = $"\"{scriptPath}\" {arguments}",
                WorkingDirectory = workingDirectory,

                UseShellExecute = true,
                CreateNoWindow = false,
                //UseShellExecute = false,
                //CreateNoWindow = true,
            };

            if (redirectOutput)
            {
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
            }

            var proc = new Process
            {
                StartInfo = psi,
                EnableRaisingEvents = true
            };

            if (redirectOutput)
            {
                proc.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        StdOutReceived?.Invoke(e.Data);
                };

                proc.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        StdErrReceived?.Invoke(e.Data);
                };
            }

            proc.Exited += (s, e) =>
            {
                int code = proc.ExitCode;
                Exited?.Invoke(code);
            };

            proc.Start();

            if (redirectOutput)
            {
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
            }

            _process = proc;
        }

        /// <summary>
        /// 要求停止 Python server（直接 Kill）。
        /// 如果希望 Python 自己優雅收尾，可以先透過 M5PythonClient 送 "quit" 再 Stop。
        /// </summary>
        public void Stop()
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.Kill(entireProcessTree: true);
                    _process.WaitForExit(2000); // 最多等 2 秒
                }
            }
            catch
            {
                // 忽略 Kill 失敗
            }
            finally
            {
                _process?.Dispose();
                _process = null;
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
