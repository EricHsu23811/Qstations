using System;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
//using System.Windows.Media.Imaging; // 這個類別不需要 WPF 相關的命名空間

namespace M5.Core
{
    /// <summary>
    /// 封裝與 Python M5 server 的 Socket 通訊：
    /// - Connect / Disconnect
    /// - 背景接收影像
    /// - 送 JSON 指令
    /// </summary>

    public enum M5ConnectionState
    {
        Disconnected,
        Connecting,
        Connected,
        Error
    }


    public class M5PythonClient : IDisposable
    {
        private TcpClient? _tcpClient;
        private NetworkStream? _stream;
        private CancellationTokenSource? _cts;
        private Task? _recvTask;

        //public bool IsConnected => _tcpClient != null && _tcpClient.Connected;
        public M5ConnectionState ConnectionState { get; private set; } = M5ConnectionState.Disconnected;
        public event Action<M5ConnectionState>? ConnectionStateChanged;

        private void SetState(M5ConnectionState newState)
        {
            if (ConnectionState == newState) return;

            ConnectionState = newState;
            ConnectionStateChanged?.Invoke(newState);
        }


        /// <summary>
        /// 當收到新的影像時觸發（在背景 thread）
        /// </summary>
        public event Action<byte[]>? FrameReceived;

        /// <summary>
        /// 發生錯誤時觸發（在背景 thread）
        /// </summary>
        public event Action<Exception>? ErrorOccurred;



        public async Task ConnectAsync(string host, int port)
        {
            if (ConnectionState == M5ConnectionState.Connected)
                return;

            SetState(M5ConnectionState.Connecting);

            try
            {
                _tcpClient = new TcpClient();
                await _tcpClient.ConnectAsync(host, port);
                _stream = _tcpClient.GetStream();

                _cts = new CancellationTokenSource();

                SetState(M5ConnectionState.Connected);

                var token = _cts.Token;

                // 啟動背景接收 Task
                _recvTask = Task.Run(() => ReceiveLoopAsync(token), token);
            }
            catch (Exception)
            {
                SetState(M5ConnectionState.Error);
                throw;
            }
        }

        public async Task DisconnectAsync(bool sendQuit = false)
        {
            try
            {
                if (sendQuit)
                {
                    await SendCommandAsync(new { cmd = "quit" });
                }
            }
            catch { /* 忽略 quit 失敗 */ }

            try
            {
                _cts?.Cancel();
                if (_recvTask != null)
                    await _recvTask;
            }
            catch { }

            try
            {
                _stream?.Close();
                _tcpClient?.Close();
            }
            catch { }

            _stream = null;
            _tcpClient = null;
            _recvTask = null;
            _cts = null;

            SetState(M5ConnectionState.Disconnected);
        }

        private async Task ReceiveLoopAsync(CancellationToken token)
        {
            byte[] lengthBuf = new byte[4];

            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (_stream == null)
                        break;

                    bool ok = await ReadExactAsync(_stream, lengthBuf, 0, 4, token);
                    if (!ok)
                        break;

                    int length =
                        lengthBuf[0] << 24 |
                        lengthBuf[1] << 16 |
                        lengthBuf[2] << 8 |
                         lengthBuf[3];

                    if (length <= 0 || length > 10_000_000)
                        break;

                    byte[] imgBuf = new byte[length];
                    ok = await ReadExactAsync(_stream, imgBuf, 0, length, token);
                    if (!ok)
                        break;

                    // 直接把原始 JPEG/影像 bytes 丟出來
                    FrameReceived?.Invoke(imgBuf);
                }
            }
            catch (OperationCanceledException)
            {
                // 正常關閉
            }
            catch (ObjectDisposedException)
            {
                // Stream 被關閉，也視為正常
            }
            catch (Exception ex)
            {
                ErrorOccurred?.Invoke(ex);
                SetState(M5ConnectionState.Error);
            }
            finally
            {
                // 連線中斷 / 離線
                SetState(M5ConnectionState.Disconnected);
            }
        }

        private static async Task<bool> ReadExactAsync(
            Stream? stream, byte[] buffer, int offset, int count, CancellationToken token)
        {
            if (stream == null)
                return false;

            int readTotal = 0;
            while (readTotal < count)
            {
                int read = await stream.ReadAsync(buffer, offset + readTotal, count - readTotal, token);
                if (read <= 0)
                    return false;

                readTotal += read;
            }
            return true;
        }

        /// <summary>
        /// 傳送任意 JSON 指令物件
        /// </summary>
        public async Task SendCommandAsync(object cmdObject)
        {
            EnsureConnected();

            string jsonLine = JsonSerializer.Serialize(cmdObject) + "\n";
            byte[] buf = Encoding.UTF8.GetBytes(jsonLine);

            await _stream.WriteAsync(buf, 0, buf.Length);
        }

        // 以下是 M5 相關的語意化方法（方便主程式呼叫）

        private void EnsureConnected()
        {
            if (ConnectionState != M5ConnectionState.Connected)
            {
                throw new InvalidOperationException("M5 is not connected.");
                // 或自訂：throw new NotConnectedException("M5 is not connected.");
            }
        }

        public Task SetLaserPowerAsync(int value) =>
            SendCommandAsync(new { cmd = "set_laser_power", value });

        public Task SetExposureAsync(int value) =>
            SendCommandAsync(new { cmd = "set_exposure", value });

        public Task SetLaserAsync(bool on) =>
            SendCommandAsync(new { cmd = "set_laser", on });

        public Task SetStreamingAsync(bool on) =>
            SendCommandAsync(new { cmd = "set_streaming", on });

        public Task SaveFrameAsync() =>
            SendCommandAsync(new { cmd = "save_frame" });

        public Task QuitAsync() =>
            SendCommandAsync(new { cmd = "quit" });

        public void Dispose()
        {
            _ = DisconnectAsync(sendQuit: false);
        }
    }
}
