using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace meiqibaojing
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            #region
            //            using CCHMIRUNTIME;
            //            using System;
            //            using System.Collections.Concurrent;
            //            using System.Collections.Generic;
            //            using System.Configuration;
            //            using System.Data;
            //            using System.Diagnostics;
            //            using System.Drawing;
            //            using System.IO.Ports;
            //            using System.Linq;
            //            using System.Text;
            //            using System.Threading;
            //            using System.Threading.Tasks;
            //            using System.Windows.Forms;

            //namespace meiqibaojing
            //    {
            //        public partial class Form1 : Form
            //        {
            //            int RESPONSE_TIMEOUT = int.Parse(ConfigurationManager.AppSettings["RESPONSETIMEOUT"]);
            //            private int[] stationAddresses; // 站地址数组
            //            private HMIRuntime hmi485 = new CCHMIRUNTIME.HMIRuntime();
            //            private bool isPaused = false;
            //            private bool stopWritingToWinCC = false;
            //            private CancellationTokenSource globalCts = new CancellationTokenSource();
            //            private List<SerialPortContext> serialPortContexts = new List<SerialPortContext>();
            //            private System.Windows.Forms.Timer twoTimer = new System.Windows.Forms.Timer();

            //            private class SerialPortContext : IDisposable
            //            {
            //                public StringBuilder AlarmMessages { get; } = new StringBuilder();
            //                public DateTime LastAlarmUpdate { get; set; } = DateTime.MinValue;
            //                public SerialPort Port { get; set; }
            //                public string Prefix { get; set; }
            //                public int StartRegisterAddress { get; set; }
            //                public int RegisterCount { get; set; }
            //                public List<byte> ReceiveBuffer { get; } = new List<byte>();
            //                public DateTime LastReceivedTime { get; set; } = DateTime.MinValue;
            //                public ConcurrentDictionary<int, TaskCompletionSource<bool>> ResponseWaiters { get; } = new ConcurrentDictionary<int, TaskCompletionSource<bool>>();
            //                public object SerialLock { get; } = new object();
            //                public bool IsPolling { get; set; } = true;
            //                public CancellationTokenSource Cts { get; } = new CancellationTokenSource();
            //                public int Index { get; set; }  // 串口索引
            //                public System.Windows.Forms.Timer FrameTimer { get; set; } // 帧超时定时器

            //                public void Dispose()
            //                {
            //                    try
            //                    {
            //                        Cts?.Cancel();
            //                        Cts?.Dispose();
            //                        FrameTimer?.Stop();
            //                        FrameTimer?.Dispose();

            //                        // 确保串口完全释放
            //                        if (Port != null)
            //                        {
            //                            try
            //                            {
            //                                if (Port.IsOpen)
            //                                    Port.Close();
            //                                Port.Dispose();
            //                                Port = null; // 设置为null防止重复使用
            //                            }
            //                            catch
            //                            {
            //                                // 忽略释放异常
            //                                Port = null;
            //                            }
            //                        }
            //                    }
            //                    catch
            //                    {
            //                        // 忽略所有异常
            //                    }
            //                }
            //            }// 串口上下文类，管理每个串口的独立状态

            //            public Form1()
            //            {
            //                InitializeComponent();
            //                LoadPorts();
            //                twoTimer.Interval = 1000;
            //                twoTimer.Tick += new EventHandler(WinccConn);
            //                twoTimer.Start();
            //            }

            //            private async void Form1_Load(object sender, EventArgs e)
            //            {
            //                Task while04 = Task.Run(() => Winccstart());
            //                try
            //                {
            //                    // 初始化所有串口
            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        try
            //                        {
            //                            context.Port.DataReceived += (s, args) => SerialPort_DataReceived(s, args, context);
            //                            context.Port.Open();
            //                            context.FrameTimer.Start(); // 启动帧超时检测定时器

            //                            // 更新WinCC状态 - 确保每个串口的状态都被更新
            //                            string tagName = $"meiqi{context.Prefix}区";
            //                            hmi485.Tags[tagName].Write(1);
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            MessageBox.Show($"初始化串口{context.Port.PortName}失败: {ex.Message}");
            //                            continue;
            //                        }
            //                    }

            //                    // 启动轮询任务
            //                    var pollingTasks = serialPortContexts
            //                        .Where(context => context.Port.IsOpen)
            //                        .Select(context => StartPollingAsync(context));

            //                    await Task.WhenAll(pollingTasks);
            //                }
            //                catch (Exception ex)
            //                {
            //                    MessageBox.Show($"初始化失败: {ex.Message}");
            //                    Environment.Exit(Environment.ExitCode);
            //                }
            //            }

            //            private void LoadPorts()
            //            {
            //                try
            //                {
            //                    int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "200");
            //                    string addressesConfig = ConfigurationManager.AppSettings["StationAddresses"];
            //                    if (!string.IsNullOrEmpty(addressesConfig))
            //                    {
            //                        stationAddresses = addressesConfig.Split(',')
            //                            .Select(addr => int.Parse(addr.Trim()))
            //                            .ToArray();
            //                    }
            //                    else
            //                    {
            //                        stationAddresses = new int[] { 1 };
            //                    }

            //                    string[] portNames = ConfigurationManager.AppSettings["SerialPorts"].Split(';');
            //                    int[] baudRates = ConfigurationManager.AppSettings["BaudRates"].Split(';').Select(int.Parse).ToArray();
            //                    Parity[] parities = ConfigurationManager.AppSettings["Parities"].Split(';').Select(p => (Parity)Enum.Parse(typeof(Parity), p)).ToArray();
            //                    int[] dataBitsArray = ConfigurationManager.AppSettings["DataBitsArray"].Split(';').Select(int.Parse).ToArray();
            //                    StopBits[] stopBitsArray = ConfigurationManager.AppSettings["StopBitsArray"].Split(';').Select(s => (StopBits)Enum.Parse(typeof(StopBits), s)).ToArray();
            //                    string[] prefixes = ConfigurationManager.AppSettings["AddressPrefixes"].Split(';');
            //                    int[] startRegisterAddresses = ConfigurationManager.AppSettings["RegisterAddresses"].Split(';').Select(int.Parse).ToArray();
            //                    int[] registerCounts;
            //                    string registerCountsConfig = ConfigurationManager.AppSettings["RegisterCounts"];
            //                    if (!string.IsNullOrEmpty(registerCountsConfig))
            //                    {
            //                        registerCounts = registerCountsConfig.Split(';').Select(int.Parse).ToArray();
            //                    }
            //                    else
            //                    {
            //                        registerCounts = new int[portNames.Length];
            //                        for (int i = 0; i < portNames.Length; i++)
            //                        {
            //                            registerCounts[i] = 16;
            //                        }
            //                    }
            //                    if (new[] { portNames.Length, baudRates.Length, parities.Length, dataBitsArray.Length, stopBitsArray.Length, prefixes.Length, startRegisterAddresses.Length, registerCounts.Length }
            //                       .Distinct().Count() != 1)
            //                    {
            //                        throw new ConfigurationErrorsException("串口配置数组长度不一致");
            //                    }

            //                    StringBuilder configText = new StringBuilder("串口配置:\r\n");
            //                    for (int i = 0; i < portNames.Length; i++)
            //                    {
            //                        configText.AppendLine($"串口{i + 1}: {portNames[i]}, {baudRates[i]}, {parities[i]}, {dataBitsArray[i]}, {stopBitsArray[i]}, 区域:{prefixes[i]}, 起始寄存器:{startRegisterAddresses[i]}, 读取数量:{registerCounts[i]}, 轮询间隔:{pollingInterval}ms");// 创建串口上下文
            //                        var context = new SerialPortContext
            //                        {
            //                            Port = new SerialPort(portNames[i], baudRates[i], parities[i], dataBitsArray[i], stopBitsArray[i]),
            //                            Prefix = prefixes[i],
            //                            StartRegisterAddress = startRegisterAddresses[i],
            //                            RegisterCount = registerCounts[i],
            //                            Index = i + 1,
            //                            FrameTimer = new System.Windows.Forms.Timer { Interval = int.Parse(ConfigurationManager.AppSettings["FrameInterval"] ?? "20") }
            //                        };

            //                        context.FrameTimer.Tick += (sender, e) => CheckFrameTimeout(context);
            //                        serialPortContexts.Add(context);
            //                    }

            //                    textBox1.Text = configText.ToString();
            //                }
            //                catch (Exception ex)
            //                {
            //                    MessageBox.Show($"加载配置失败: {ex.Message}");
            //                    Environment.Exit(1);
            //                }
            //            }

            //            private void WinccConn(object sender, EventArgs e)
            //            {
            //                try
            //                {
            //                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;

            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        try
            //                        {
            //                            // 添加 null 检查
            //                            if (context == null || context.Port == null)
            //                            {
            //                                continue;
            //                            }

            //                            if (context.Port.IsOpen)
            //                            {
            //                                string tagName = $"meiqi{context.Prefix}区";
            //                                hmi485.Tags[tagName]?.Write(1);
            //                            }
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"更新串口状态到WinCC失败: {ex.Message}");
            //                            // 即使出错也继续处理其他串口
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;
            //                    MessageBox.Show($"与WinCC通讯失败: {ex.Message}，驱动程序将退出。");
            //                    Environment.Exit(Environment.ExitCode);
            //                }
            //            }

            //            private async Task Winccstart()
            //            {
            //                string processName = ConfigurationManager.AppSettings["ProcessName"];

            //                while (!globalCts.IsCancellationRequested)
            //                {
            //                    try
            //                    {
            //                        bool b = IsProcessStarted(processName);

            //                        if (b)
            //                        {
            //                            stopWritingToWinCC = false;
            //                            // 更新所有串口的状态
            //                            foreach (var context in serialPortContexts)
            //                            {
            //                                try
            //                                {
            //                                    // 添加 null 检查
            //                                    if (context == null || context.Port == null)
            //                                    {
            //                                        continue;
            //                                    }

            //                                    if (context.Port.IsOpen)
            //                                    {
            //                                        string tagName = $"meiqi{context.Prefix}区";
            //                                        hmi485.Tags[tagName]?.Write(1);
            //                                    }
            //                                }
            //                                catch (Exception ex)
            //                                {
            //                                    Debug.WriteLine($"更新串口{context?.Port?.PortName}状态到WinCC失败: {ex.Message}");
            //                                }
            //                            }
            //                        }
            //                        else
            //                        {
            //                            if (InvokeRequired)
            //                            {
            //                                Invoke(new Action(() =>
            //                                {
            //                                    DialogResult result = MessageBox.Show(
            //                                        $"未找到运行的WinCC实例({processName})，是否退出驱动程序？",
            //                                        "WinCC未运行",
            //                                        MessageBoxButtons.YesNo,
            //                                        MessageBoxIcon.Warning);

            //                                    if (result == DialogResult.Yes)
            //                                    {
            //                                        MessageBox.Show("驱动程序将退出。");
            //                                        Environment.Exit(Environment.ExitCode);
            //                                    }
            //                                    else
            //                                    {
            //                                        stopWritingToWinCC = true;
            //                                        richTextBox1.AppendText($"已停止向WinCC写入标签。WinCC进程({processName})未运行。\r\n");
            //                                        richTextBox1.ScrollToCaret();
            //                                    }
            //                                }));
            //                            }

            //                            await Task.Delay(10000, globalCts.Token);
            //                        }

            //                        await Task.Delay(1000, globalCts.Token);
            //                    }
            //                    catch (OperationCanceledException)
            //                    {
            //                        // 任务被取消，正常退出
            //                        return;
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"WinCC进程检测错误: {ex.Message}");
            //                        await Task.Delay(5000, globalCts.Token);
            //                    }
            //                }
            //            }

            //            bool IsProcessStarted(string processName1)
            //            {
            //                try
            //                {
            //                    Process[] temp = Process.GetProcessesByName(processName1);
            //                    return temp.Length > 0;
            //                }
            //                catch
            //                {
            //                    return false;
            //                }
            //            }

            //            private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e, SerialPortContext context)
            //            {
            //                try
            //                {
            //                    lock (context.SerialLock)
            //                    {
            //                        int bytesToRead = context.Port.BytesToRead;
            //                        if (bytesToRead <= 0) return;

            //                        byte[] buffer = new byte[bytesToRead];
            //                        context.Port.Read(buffer, 0, bytesToRead);

            //                        // 更新最后接收时间
            //                        context.LastReceivedTime = DateTime.Now;

            //                        // 将数据添加到缓冲区
            //                        context.ReceiveBuffer.AddRange(buffer);

            //                        // 重置帧超时定时器
            //                        context.FrameTimer.Stop();
            //                        context.FrameTimer.Start();

            //                        // 尝试处理缓冲区中的完整帧
            //                        ProcessReceivedData(context);
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"SerialPort_DataReceived error: {ex.Message}");
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText($"串口{context.Port.PortName}接收数据错误: {ex.Message}\r\n");
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }

            //            private void ProcessReceivedData(SerialPortContext context)
            //            {
            //                try
            //                {
            //                    if (context.ReceiveBuffer.Count == 0) return;

            //                    // 查找可能的完整帧
            //                    while (context.ReceiveBuffer.Count >= 5) // 最小帧长度
            //                    {
            //                        // 检查是否为错误响应帧（5字节）
            //                        if (context.ReceiveBuffer.Count >= 5 &&
            //                            (context.ReceiveBuffer[1] & 0x80) != 0)
            //                        {
            //                            // 错误响应帧长度为5字节
            //                            if (context.ReceiveBuffer.Count >= 5)
            //                            {
            //                                byte[] errorFrame = context.ReceiveBuffer.Take(5).ToArray();
            //                                ProcessCompleteFrame(errorFrame, context);
            //                                context.ReceiveBuffer.RemoveRange(0, 5);
            //                                continue;
            //                            }
            //                        }

            //                        // 检查正常响应帧
            //                        if (context.ReceiveBuffer[1] == 0x03) // 读取保持寄存器
            //                        {
            //                            // 正常响应帧长度 = 地址(1) + 功能码(1) + 字节数(1) + 数据(n*2) + CRC(2)
            //                            int expectedLength = 3 + context.ReceiveBuffer[2] + 2;

            //                            if (context.ReceiveBuffer.Count >= expectedLength)
            //                            {
            //                                byte[] completeFrame = context.ReceiveBuffer.Take(expectedLength).ToArray();
            //                                ProcessCompleteFrame(completeFrame, context);
            //                                context.ReceiveBuffer.RemoveRange(0, expectedLength);
            //                                continue;
            //                            }
            //                            else
            //                            {
            //                                // 帧不完整，等待更多数据
            //                                break;
            //                            }
            //                        }
            //                        else
            //                        {
            //                            // 未知帧类型，尝试查找下一个可能的帧头
            //                            int nextStart = FindNextFrameStart(context.ReceiveBuffer);
            //                            if (nextStart > 0)
            //                            {
            //                                // 移除无效数据
            //                                context.ReceiveBuffer.RemoveRange(0, nextStart);
            //                                continue;
            //                            }
            //                            else
            //                            {
            //                                // 没有找到有效帧头，清空缓冲区
            //                                context.ReceiveBuffer.Clear();
            //                                break;
            //                            }
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"ProcessReceivedData error: {ex.Message}");
            //                    context.ReceiveBuffer.Clear();
            //                }
            //            }

            //            private int FindNextFrameStart(List<byte> buffer)
            //            {
            //                for (int i = 1; i < buffer.Count; i++)
            //                {
            //                    // 有效的Modbus地址范围是1-247
            //                    if (buffer[i] >= 1 && buffer[i] <= 247)
            //                    {
            //                        // 检查是否是有效的功能码
            //                        if (i + 1 < buffer.Count)
            //                        {
            //                            byte functionCode = buffer[i + 1];
            //                            if (functionCode == 0x03 || (functionCode & 0x80) != 0)
            //                            {
            //                                return i;
            //                            }
            //                        }
            //                    }
            //                }
            //                return -1;
            //            }

            //            // 处理完整的帧
            //            private void ProcessCompleteFrame(byte[] frame, SerialPortContext context)
            //            {
            //                try
            //                {
            //                    // 显示接收到的原始数据
            //                    string rawDataHex = BitConverter.ToString(frame).Replace("-", " ");
            //                    string displayText = $"{context.Prefix}区（{context.Port.PortName}）收到{frame.Length}字节: {rawDataHex} -> {DateTime.Now:HH:mm:ss.fff} \r\n";

            //                    // 处理Modbus响应
            //                    bool isModbusResponse = ProcessModbusResponse(frame, context);

            //                    if (!isModbusResponse)
            //                    {
            //                        displayText += $"-> 非标准Modbus响应或CRC校验失败\r\n";
            //                    }

            //                    // 更新UI显示
            //                    if (!isPaused && !globalCts.IsCancellationRequested)
            //                    {
            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText(displayText);
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"ProcessCompleteFrame error: {ex.Message}");
            //                }
            //            }

            //            // 帧超时检测
            //            private void CheckFrameTimeout(SerialPortContext context)
            //            {
            //                lock (context.SerialLock)
            //                {
            //                    context.FrameTimer.Stop();

            //                    if (context.ReceiveBuffer.Count > 0)
            //                    {
            //                        // 帧超时，处理缓冲区中可能的不完整数据
            //                        string timeoutText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 帧超时，丢弃{context.ReceiveBuffer.Count}字节不完整数据\r\n";

            //                        if (!isPaused && !globalCts.IsCancellationRequested)
            //                        {
            //                            this.BeginInvoke(new Action(() =>
            //                            {
            //                                richTextBox1.AppendText(timeoutText);
            //                                richTextBox1.ScrollToCaret();
            //                            }));
            //                        }

            //                        context.ReceiveBuffer.Clear();
            //                    }

            //                    context.FrameTimer.Start();
            //                }
            //            }

            //            // 在 StartPollingAsync 方法中修改串口状态检查逻辑
            //            private async Task StartPollingAsync(SerialPortContext context)
            //            {
            //                int currentIndex = 0;
            //                int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "100");

            //                while (context.IsPolling && !context.Cts.Token.IsCancellationRequested && !globalCts.IsCancellationRequested)
            //                {
            //                    if (stationAddresses == null || stationAddresses.Length == 0)
            //                    {
            //                        await Task.Delay(100, context.Cts.Token);
            //                        continue;
            //                    }

            //                    int address = stationAddresses[currentIndex];
            //                    currentIndex = (currentIndex + 1) % stationAddresses.Length;

            //                    try
            //                    {
            //                        // 发送请求
            //                        byte[] request = BuildModbusRequest(address, context.StartRegisterAddress, context.RegisterCount, context);

            //                        // 获取该地址的等待器
            //                        var tcs = new TaskCompletionSource<bool>();
            //                        context.ResponseWaiters[address] = tcs;

            //                        using (var ctsTimeout = new CancellationTokenSource(RESPONSE_TIMEOUT))
            //                        {
            //                            var responseTask = tcs.Task;
            //                            var timeoutTask = Task.Delay(Timeout.Infinite, ctsTimeout.Token);

            //                            var completedTask = await Task.WhenAny(responseTask, timeoutTask);

            //                            if (completedTask == responseTask)
            //                            {
            //                                // 收到响应，取消超时
            //                                ctsTimeout.Cancel();
            //                            }
            //                            else
            //                            {
            //                                // 超时处理
            //                                context.ResponseWaiters.TryRemove(address, out _);
            //                                HandleResponseTimeout(context, address);
            //                            }
            //                        }
            //                    }
            //                    catch (OperationCanceledException)
            //                    {
            //                        // 取消操作，继续下一个
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"轮询错误: {ex.Message}");
            //                        await Task.Delay(1000, context.Cts.Token);
            //                    }

            //                    // 添加轮询间隔，避免过于频繁
            //                    await Task.Delay(pollingInterval, context.Cts.Token);
            //                }
            //            }

            //            // 处理响应超时
            //            private void HandleResponseTimeout(SerialPortContext context, int address)
            //            {
            //                if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                {
            //                    try
            //                    {
            //                        // 清空并设置报警消息
            //                        context.AlarmMessages.Clear();
            //                        string timeoutMessage = $"{context.Prefix}区({context.Port.PortName}) 报警器分线箱响应超时";
            //                        context.AlarmMessages.AppendLine(timeoutMessage);

            //                        // 更新报警消息标签 - 确保使用正确的标签名
            //                        string messageTag = $"meiqi{context.Prefix}";
            //                        hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());

            //                        // 更新所有寄存器的超时状态
            //                        for (int i = 1; i <= context.RegisterCount; i++)
            //                        {
            //                            hmi485.Tags[$"{context.Prefix}{i}_Value"].Write(999999); // 使用特殊值表示超时
            //                            hmi485.Tags[$"{context.Prefix}{i}_Time"].Write(DateTime.Now.ToString("离线时间：yyyy-MM-dd HH:mm:ss.fff"));
            //                        }

            //                        // 记录超时日志
            //                        Debug.WriteLine($"地址 {address} 响应超时，已更新WinCC标签");

            //                        if (!isPaused && !globalCts.IsCancellationRequested)
            //                        {
            //                            this.BeginInvoke(new Action(() =>
            //                            {
            //                                richTextBox1.AppendText($"{DateTime.Now:HH:mm:ss.fff} {context.Prefix}区{context.Port.PortName}: 报警器分线箱响应超时，已更新WinCC标签\r\n");
            //                                richTextBox1.ScrollToCaret();
            //                            }));
            //                        }
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"写入WinCC超时状态失败: {ex.Message}");

            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText($"写入WinCC超时状态失败: {ex.Message}\r\n");
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }
            //                }
            //            }

            //            private byte[] BuildModbusRequest(int address, int startRegister, int registerCount, SerialPortContext context)
            //            {
            //                byte[] request = new byte[8];
            //                request[0] = (byte)address;          // 设备地址
            //                request[1] = 0x03;                   // 功能码
            //                request[2] = (byte)(startRegister >> 8);  // 寄存器地址高字节
            //                request[3] = (byte)(startRegister & 0xFF);// 寄存器地址低字节
            //                request[4] = (byte)(registerCount >> 8);  // 读取数量高字节
            //                request[5] = (byte)(registerCount & 0xFF);// 读取数量低字节

            //                // 计算CRC校验
            //                byte[] crc;
            //                CRC_16(request.Take(6).ToArray(), out crc);
            //                request[6] = crc[0];
            //                request[7] = crc[1];

            //                // 将字节数组转换为十六进制字符串
            //                string hexString = BitConverter.ToString(request).Replace("-", " ");

            //                // 显示完整的请求帧
            //                if (!isPaused && !globalCts.IsCancellationRequested)
            //                {
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText($"{context.Prefix}区{context.Port.PortName}请求: {hexString}\r\n");
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }

            //                // 实际发送数据
            //                lock (context.SerialLock)
            //                {
            //                    if (context.Port.IsOpen && !globalCts.IsCancellationRequested)
            //                    {
            //                        try
            //                        {
            //                            context.Port.Write(request, 0, request.Length);
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"发送数据失败: {ex.Message}");
            //                        }
            //                    }
            //                }

            //                return request;
            //            }

            //            private bool ProcessModbusResponse(byte[] response, SerialPortContext context)
            //            {
            //                if (response.Length < 5) return false;

            //                try
            //                {
            //                    // 在处理响应前清空报警消息
            //                    context.AlarmMessages.Clear();

            //                    // 验证CRC
            //                    byte[] crc;
            //                    CRC_16(response.Take(response.Length - 2).ToArray(), out crc);
            //                    if (crc[0] != response[response.Length - 2] || crc[1] != response[response.Length - 1])
            //                    {
            //                        Debug.WriteLine("CRC校验失败");
            //                        return false;
            //                    }

            //                    int address = response[0]; // 设备地址
            //                    int functionCode = response[1];

            //                    // 检查错误响应
            //                    if ((functionCode & 0x80) != 0)
            //                    {
            //                        int errorCode = response[2];
            //                        string errorMessage = GetModbusErrorDescription(errorCode);
            //                        string errorText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 地址{address}错误响应: {errorMessage}\r\n";

            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText(errorText);
            //                            richTextBox1.ScrollToCaret();
            //                        }));

            //                        return true;
            //                    }

            //                    if (functionCode == 0x03) // 读取保持寄存器响应
            //                    {
            //                        int byteCount = response[2];
            //                        if (byteCount != context.RegisterCount * 2) return false;

            //                        // 通知等待器
            //                        if (context.ResponseWaiters.TryRemove(address, out var tcs))
            //                        {
            //                            tcs.TrySetResult(true);
            //                        }

            //                        // 解析所有寄存器值并更新到WinCC
            //                        for (int i = 0; i < context.RegisterCount; i++)
            //                        {
            //                            int registerIndex = 3 + i * 2; // 数据起始位置
            //                            int value = (response[registerIndex] << 8) | response[registerIndex + 1];
            //                            int registerNumber = i + 1; // 寄存器编号从1开始

            //                            // 处理特殊值和更新WinCC
            //                            ProcessRegisterValue(address, registerNumber, value, context);
            //                        }

            //                        // 在处理完所有寄存器后更新报警消息
            //                        if (context.AlarmMessages.Length > 0)
            //                        {
            //                            try
            //                            {
            //                                string messageTag = $"meiqi{context.Prefix}";
            //                                hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());
            //                                context.LastAlarmUpdate = DateTime.Now;
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                Debug.WriteLine($"更新报警消息失败: {ex.Message}");
            //                            }
            //                        }
            //                        else
            //                        {
            //                            // 如果没有报警，清空报警消息
            //                            try
            //                            {
            //                                string messageTag = $"meiqi{context.Prefix}";
            //                                hmi485.Tags[messageTag].Write("无报警");
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                Debug.WriteLine($"清空报警消息失败: {ex.Message}");
            //                            }
            //                        }

            //                        return true;
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}区响应，报警消息长度: {context.AlarmMessages.Length}");
            //                    Debug.WriteLine($"ProcessModbusResponse error: {ex.Message}");
            //                }

            //                return false;
            //            }

            //            private void ProcessRegisterValue(int address, int registerNumber, int value, SerialPortContext context)
            //            {
            //                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //                string displayText = $"{context.Prefix}区（{context.Port.PortName}）第{registerNumber}路数据: 值：{value} 读取时间：{timestamp}，";
            //                string messageTag = $"meiqi{context.Prefix}";
            //                try
            //                {
            //                    if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                    {
            //                        string statusTag = $"meiqi{context.Prefix}区";
            //                        string valueTag = $"{context.Prefix}{registerNumber}_Value";
            //                        string timeTag = $"{context.Prefix}{registerNumber}_Time";

            //                        hmi485.Tags[statusTag].Write(1);

            //                        // 处理特殊值
            //                        if (value == 0xFAFA) // 探测器没有接
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFAFA);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 没有连接");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器没有连接；");
            //                            displayText += $"-> 探测器没有连接\r\n";
            //                        }
            //                        else if (value == 0xFEFE) // 探测器丢失
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFEFE);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 丢失");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器丢失");
            //                            displayText += $"-> 探测器丢失\r\n";
            //                        }
            //                        else if (value == 0xFCFC) // 探测器故障
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFCFC);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 故障");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器故障");
            //                            displayText += $"-> 探测器故障\r\n";
            //                        }
            //                        else // 正常值
            //                        {
            //                            float actualValue = value / 10.0f;
            //                            hmi485.Tags[valueTag].Write(actualValue);
            //                            hmi485.Tags[timeTag].Write(timestamp);
            //                            hmi485.Tags[messageTag].Write($"CO报警地址 {context.Prefix}{registerNumber}  值: {actualValue:F1}ppm");
            //                            // 正常值不添加到报警消息中
            //                            string tagName = $"meiqi{context.Prefix}区";
            //                            hmi485.Tags[tagName].Write(1);
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}{registerNumber}，值: 0x{value:X4}");
            //                    displayText += $"! 写入WinCC标签失败: {ex.Message}\r\n";
            //                }

            //                // 更新UI
            //                if (!isPaused && !globalCts.IsCancellationRequested)
            //                {
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText(displayText);
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }

            //            private string GetModbusErrorDescription(int errorCode)
            //            {
            //                switch (errorCode)
            //                {
            //                    case 0x01: return "非法功能";
            //                    case 0x02: return "非法数据地址";
            //                    case 0x03: return "非法数据值";
            //                    case 0x04: return "从站设备故障";
            //                    case 0x05: return "确认";
            //                    case 0x06: return "从属设备忙";
            //                    case 0x08: return "存储奇偶性差错";
            //                    case 0x0A: return "不可用网关路径";
            //                    case 0x0B: return "网关目标设备响应失败";
            //                    default: return "未知错误";
            //                }
            //            }

            //            public static byte[] CRC_16(byte[] data, out byte[] temdata)
            //            {
            //                if (data.Length == 0)
            //                    throw new Exception("调用CRC16校验算法,（低字节在前，高字节在后）时发生异常，异常信息：被校验的数组长度为0。");
            //                int xda, xdapoly;
            //                byte i, j, xdabit;
            //                xda = 0xFFFF;
            //                xdapoly = 0xA001;
            //                for (i = 0; i < data.Length; i++)
            //                {
            //                    xda ^= data[i];
            //                    for (j = 0; j < 8; j++)
            //                    {
            //                        xdabit = (byte)(xda & 0x01);
            //                        xda >>= 1;
            //                        if (xdabit == 1)
            //                            xda ^= xdapoly;
            //                    }
            //                }
            //                temdata = new byte[2] { (byte)(xda & 0xFF), (byte)(xda >> 8) };
            //                return temdata;
            //            }

            //            private void Form1_FormClosing(object sender, FormClosingEventArgs e)
            //            {
            //                // 停止所有后台操作
            //                globalCts.Cancel();
            //                twoTimer.Stop();

            //                // 停止所有串口轮询并关闭串口
            //                foreach (var context in serialPortContexts)
            //                {
            //                    context.IsPolling = false;
            //                    context.Cts.Cancel();
            //                    context.FrameTimer.Stop();
            //                    if (context.Port.IsOpen) context.Port.Close();
            //                    context.Dispose();
            //                }

            //                // 更新WinCC状态为关闭 - 确保所有前缀的状态都被更新
            //                try
            //                {
            //                    if (!stopWritingToWinCC)
            //                    {
            //                        foreach (var context in serialPortContexts)
            //                        {
            //                            try
            //                            {
            //                                //string tagName = $"meiqi{context.Prefix}区";
            //                                //hmi485.Tags[tagName].Write(0);
            //                            }
            //                            catch { /* 忽略单个串口关闭时的异常 */ }
            //                        }
            //                    }
            //                }
            //                catch { /* 忽略关闭时的异常 */ }
            //            }

            //            private void richTextBox1_TextChanged(object sender, EventArgs e)
            //            {
            //                if (isPaused || globalCts.IsCancellationRequested) return;

            //                const int maxLine = 1200;
            //                if (richTextBox1.Lines.Length > maxLine)
            //                {
            //                    int firstNewLine = richTextBox1.Text.IndexOf('\n');
            //                    if (firstNewLine >= 0)
            //                    {
            //                        // 保留滚动位置
            //                        int scrollPos = richTextBox1.GetCharIndexFromPosition(new Point(0, 0));
            //                        richTextBox1.Text = richTextBox1.Text.Substring(firstNewLine + 1);
            //                        richTextBox1.ScrollToCaret();
            //                        if (scrollPos > 0)
            //                        {
            //                            richTextBox1.SelectionStart = Math.Min(scrollPos - firstNewLine - 1, richTextBox1.Text.Length);
            //                            richTextBox1.ScrollToCaret();
            //                        }
            //                    }
            //                }
            //            }

            //            protected override void Dispose(bool disposing)
            //            {
            //                if (disposing)
            //                {
            //                    try
            //                    {
            //                        // 取消所有操作
            //                        globalCts?.Cancel();

            //                        // 停止并释放定时器
            //                        twoTimer?.Stop();
            //                        twoTimer?.Dispose();

            //                        // 释放所有串口上下文
            //                        foreach (var context in serialPortContexts)
            //                        {
            //                            context?.Dispose();
            //                        }
            //                        serialPortContexts.Clear();

            //                        // 释放组件资源
            //                        components?.Dispose();
            //                    }
            //                    catch
            //                    {
            //                        // 忽略清理过程中的异常
            //                    }
            //                    finally
            //                    {
            //                        globalCts?.Dispose();
            //                    }
            //                }
            //                base.Dispose(disposing);
            //            }

            //            private void button1_Click(object sender, EventArgs e) { isPaused = !isPaused; button1.Text = isPaused ? "继续滚屏显示" : "停止滚屏显示"; }
            //            private void textBox1_TextChanged(object sender, EventArgs e) { }
            //        }
            //    }

            #endregion
            #region//不再更新掉线的通讯状态
            //            using CCHMIRUNTIME;
            //            using System;
            //            using System.Collections.Concurrent;
            //            using System.Collections.Generic;
            //            using System.Configuration;
            //            using System.Data;
            //            using System.Diagnostics;
            //            using System.Drawing;
            //            using System.IO.Ports;
            //            using System.Linq;
            //            using System.Text;
            //            using System.Threading;
            //            using System.Threading.Tasks;
            //            using System.Windows.Forms;

            //namespace meiqibaojing
            //    {
            //        public partial class Form1 : Form
            //        {
            //            int RESPONSE_TIMEOUT = int.Parse(ConfigurationManager.AppSettings["RESPONSETIMEOUT"]);
            //            private int[] stationAddresses; // 站地址数组
            //            private HMIRuntime hmi485 = new CCHMIRUNTIME.HMIRuntime();
            //            private bool isPaused = false;
            //            private bool stopWritingToWinCC = false;
            //            private CancellationTokenSource globalCts = new CancellationTokenSource();
            //            private List<SerialPortContext> serialPortContexts = new List<SerialPortContext>();
            //            private System.Windows.Forms.Timer twoTimer = new System.Windows.Forms.Timer();
            //            // 添加超时状态字典
            //            private ConcurrentDictionary<string, bool> portTimeoutStatus = new ConcurrentDictionary<string, bool>();

            //            private class SerialPortContext : IDisposable
            //            {
            //                public StringBuilder AlarmMessages { get; } = new StringBuilder();
            //                public DateTime LastAlarmUpdate { get; set; } = DateTime.MinValue;
            //                public SerialPort Port { get; set; }
            //                public string Prefix { get; set; }
            //                public int StartRegisterAddress { get; set; }
            //                public int RegisterCount { get; set; }
            //                public List<byte> ReceiveBuffer { get; } = new List<byte>();
            //                public DateTime LastReceivedTime { get; set; } = DateTime.MinValue;
            //                public ConcurrentDictionary<int, TaskCompletionSource<bool>> ResponseWaiters { get; } = new ConcurrentDictionary<int, TaskCompletionSource<bool>>();
            //                public object SerialLock { get; } = new object();
            //                public bool IsPolling { get; set; } = true;
            //                public CancellationTokenSource Cts { get; } = new CancellationTokenSource();
            //                public int Index { get; set; }  // 串口索引
            //                public System.Windows.Forms.Timer FrameTimer { get; set; } // 帧超时定时器
            //                public bool IsTimeout { get; set; } = false; // 添加超时状态

            //                public void Dispose()
            //                {
            //                    try
            //                    {
            //                        Cts?.Cancel();
            //                        Cts?.Dispose();
            //                        FrameTimer?.Stop();
            //                        FrameTimer?.Dispose();

            //                        // 确保串口完全释放
            //                        if (Port != null)
            //                        {
            //                            try
            //                            {
            //                                if (Port.IsOpen)
            //                                    Port.Close();
            //                                Port.Dispose();
            //                                Port = null; // 设置为null防止重复使用
            //                            }
            //                            catch
            //                            {
            //                                // 忽略释放异常
            //                                Port = null;
            //                            }
            //                        }
            //                    }
            //                    catch
            //                    {
            //                        // 忽略所有异常
            //                    }
            //                }
            //            }// 串口上下文类，管理每个串口的独立状态

            //            public Form1()
            //            {
            //                InitializeComponent();
            //                LoadPorts();
            //                twoTimer.Interval = 1000;
            //                twoTimer.Tick += new EventHandler(WinccConn);
            //                twoTimer.Start();
            //            }

            //            private async void Form1_Load(object sender, EventArgs e)
            //            {
            //                Task while04 = Task.Run(() => Winccstart());
            //                try
            //                {
            //                    // 初始化所有串口
            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        try
            //                        {
            //                            context.Port.DataReceived += (s, args) => SerialPort_DataReceived(s, args, context);
            //                            context.Port.Open();
            //                            context.FrameTimer.Start(); // 启动帧超时检测定时器

            //                            // 更新WinCC状态 - 确保每个串口的状态都被更新
            //                            string tagName = $"meiqi{context.Prefix}区";
            //                            hmi485.Tags[tagName].Write(1);
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            MessageBox.Show($"初始化串口{context.Port.PortName}失败: {ex.Message}");
            //                            continue;
            //                        }
            //                    }

            //                    // 启动轮询任务
            //                    var pollingTasks = serialPortContexts
            //                        .Where(context => context.Port.IsOpen)
            //                        .Select(context => StartPollingAsync(context));

            //                    await Task.WhenAll(pollingTasks);
            //                }
            //                catch (Exception ex)
            //                {
            //                    MessageBox.Show($"初始化失败: {ex.Message}");
            //                    Environment.Exit(Environment.ExitCode);
            //                }
            //            }

            //            private void LoadPorts()
            //            {
            //                try
            //                {
            //                    int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "200");
            //                    string addressesConfig = ConfigurationManager.AppSettings["StationAddresses"];
            //                    if (!string.IsNullOrEmpty(addressesConfig))
            //                    {
            //                        stationAddresses = addressesConfig.Split(',')
            //                            .Select(addr => int.Parse(addr.Trim()))
            //                            .ToArray();
            //                    }
            //                    else
            //                    {
            //                        stationAddresses = new int[] { 1 };
            //                    }

            //                    string[] portNames = ConfigurationManager.AppSettings["SerialPorts"].Split(';');
            //                    int[] baudRates = ConfigurationManager.AppSettings["BaudRates"].Split(';').Select(int.Parse).ToArray();
            //                    Parity[] parities = ConfigurationManager.AppSettings["Parities"].Split(';').Select(p => (Parity)Enum.Parse(typeof(Parity), p)).ToArray();
            //                    int[] dataBitsArray = ConfigurationManager.AppSettings["DataBitsArray"].Split(';').Select(int.Parse).ToArray();
            //                    StopBits[] stopBitsArray = ConfigurationManager.AppSettings["StopBitsArray"].Split(';').Select(s => (StopBits)Enum.Parse(typeof(StopBits), s)).ToArray();
            //                    string[] prefixes = ConfigurationManager.AppSettings["AddressPrefixes"].Split(';');
            //                    int[] startRegisterAddresses = ConfigurationManager.AppSettings["RegisterAddresses"].Split(';').Select(int.Parse).ToArray();
            //                    int[] registerCounts;
            //                    string registerCountsConfig = ConfigurationManager.AppSettings["RegisterCounts"];
            //                    if (!string.IsNullOrEmpty(registerCountsConfig))
            //                    {
            //                        registerCounts = registerCountsConfig.Split(';').Select(int.Parse).ToArray();
            //                    }
            //                    else
            //                    {
            //                        registerCounts = new int[portNames.Length];
            //                        for (int i = 0; i < portNames.Length; i++)
            //                        {
            //                            registerCounts[i] = 16;
            //                        }
            //                    }
            //                    if (new[] { portNames.Length, baudRates.Length, parities.Length, dataBitsArray.Length, stopBitsArray.Length, prefixes.Length, startRegisterAddresses.Length, registerCounts.Length }
            //                       .Distinct().Count() != 1)
            //                    {
            //                        throw new ConfigurationErrorsException("串口配置数组长度不一致");
            //                    }

            //                    StringBuilder configText = new StringBuilder("串口配置:\r\n");
            //                    for (int i = 0; i < portNames.Length; i++)
            //                    {
            //                        configText.AppendLine($"串口{i + 1}: {portNames[i]}, {baudRates[i]}, {parities[i]}, {dataBitsArray[i]}, {stopBitsArray[i]}, 区域:{prefixes[i]}, 起始寄存器:{startRegisterAddresses[i]}, 读取数量:{registerCounts[i]}, 轮询间隔:{pollingInterval}ms");// 创建串口上下文
            //                        var context = new SerialPortContext
            //                        {
            //                            Port = new SerialPort(portNames[i], baudRates[i], parities[i], dataBitsArray[i], stopBitsArray[i]),
            //                            Prefix = prefixes[i],
            //                            StartRegisterAddress = startRegisterAddresses[i],
            //                            RegisterCount = registerCounts[i],
            //                            Index = i + 1,
            //                            FrameTimer = new System.Windows.Forms.Timer { Interval = int.Parse(ConfigurationManager.AppSettings["FrameInterval"] ?? "20") }
            //                        };

            //                        context.FrameTimer.Tick += (sender, e) => CheckFrameTimeout(context);
            //                        serialPortContexts.Add(context);
            //                        // 初始化超时状态
            //                        portTimeoutStatus[context.Prefix] = false;
            //                    }

            //                    textBox1.Text = configText.ToString();
            //                }
            //                catch (Exception ex)
            //                {
            //                    MessageBox.Show($"加载配置失败: {ex.Message}");
            //                    Environment.Exit(1);
            //                }
            //            }

            //            private void WinccConn(object sender, EventArgs e)
            //            {
            //                try
            //                {
            //                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;

            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        try
            //                        {
            //                            // 添加 null 检查
            //                            if (context == null || context.Port == null)
            //                            {
            //                                continue;
            //                            }

            //                            // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                            if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                            {
            //                                continue;
            //                            }

            //                            if (context.Port.IsOpen)
            //                            {
            //                                string tagName = $"meiqi{context.Prefix}区";
            //                                hmi485.Tags[tagName]?.Write(1);
            //                            }
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"更新串口状态到WinCC失败: {ex.Message}");
            //                            // 即使出错也继续处理其他串口
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;
            //                    MessageBox.Show($"与WinCC通讯失败: {ex.Message}，驱动程序将退出。");
            //                    Environment.Exit(Environment.ExitCode);
            //                }
            //            }

            //            private async Task Winccstart()
            //            {
            //                string processName = ConfigurationManager.AppSettings["ProcessName"];

            //                while (!globalCts.IsCancellationRequested)
            //                {
            //                    try
            //                    {
            //                        bool b = IsProcessStarted(processName);

            //                        if (b)
            //                        {
            //                            stopWritingToWinCC = false;
            //                            // 更新所有串口的状态
            //                            foreach (var context in serialPortContexts)
            //                            {
            //                                try
            //                                {
            //                                    // 添加 null 检查
            //                                    if (context == null || context.Port == null)
            //                                    {
            //                                        continue;
            //                                    }

            //                                    // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                                    if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                                    {
            //                                        continue;
            //                                    }

            //                                    if (context.Port.IsOpen)
            //                                    {
            //                                        string tagName = $"meiqi{context.Prefix}区";
            //                                        hmi485.Tags[tagName]?.Write(1);
            //                                    }
            //                                }
            //                                catch (Exception ex)
            //                                {
            //                                    Debug.WriteLine($"更新串口{context?.Port?.PortName}状态到WinCC失败: {ex.Message}");
            //                                }
            //                            }
            //                        }
            //                        else
            //                        {
            //                            if (InvokeRequired)
            //                            {
            //                                Invoke(new Action(() =>
            //                                {
            //                                    DialogResult result = MessageBox.Show(
            //                                        $"未找到运行的WinCC实例({processName})，是否退出驱动程序？",
            //                                        "WinCC未运行",
            //                                        MessageBoxButtons.YesNo,
            //                                        MessageBoxIcon.Warning);

            //                                    if (result == DialogResult.Yes)
            //                                    {
            //                                        MessageBox.Show("驱动程序将退出。");
            //                                        Environment.Exit(Environment.ExitCode);
            //                                    }
            //                                    else
            //                                    {
            //                                        stopWritingToWinCC = true;
            //                                        richTextBox1.AppendText($"已停止向WinCC写入标签。WinCC进程({processName})未运行。\r\n");
            //                                        richTextBox1.ScrollToCaret();
            //                                    }
            //                                }));
            //                            }

            //                            await Task.Delay(10000, globalCts.Token);
            //                        }

            //                        await Task.Delay(1000, globalCts.Token);
            //                    }
            //                    catch (OperationCanceledException)
            //                    {
            //                        // 任务被取消，正常退出
            //                        return;
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"WinCC进程检测错误: {ex.Message}");
            //                        await Task.Delay(5000, globalCts.Token);
            //                    }
            //                }
            //            }

            //            bool IsProcessStarted(string processName1)
            //            {
            //                try
            //                {
            //                    Process[] temp = Process.GetProcessesByName(processName1);
            //                    return temp.Length > 0;
            //                }
            //                catch
            //                {
            //                    return false;
            //                }
            //            }

            //            private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e, SerialPortContext context)
            //            {
            //                try
            //                {
            //                    lock (context.SerialLock)
            //                    {
            //                        int bytesToRead = context.Port.BytesToRead;
            //                        if (bytesToRead <= 0) return;

            //                        byte[] buffer = new byte[bytesToRead];
            //                        context.Port.Read(buffer, 0, bytesToRead);

            //                        // 更新最后接收时间
            //                        context.LastReceivedTime = DateTime.Now;

            //                        // 将数据添加到缓冲区
            //                        context.ReceiveBuffer.AddRange(buffer);

            //                        // 重置帧超时定时器
            //                        context.FrameTimer.Stop();
            //                        context.FrameTimer.Start();

            //                        // 尝试处理缓冲区中的完整帧
            //                        ProcessReceivedData(context);
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"SerialPort_DataReceived error: {ex.Message}");
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText($"串口{context.Port.PortName}接收数据错误: {ex.Message}\r\n");
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }

            //            private void ProcessReceivedData(SerialPortContext context)
            //            {
            //                try
            //                {
            //                    if (context.ReceiveBuffer.Count == 0) return;

            //                    // 查找可能的完整帧
            //                    while (context.ReceiveBuffer.Count >= 5) // 最小帧长度
            //                    {
            //                        // 检查是否为错误响应帧（5字节）
            //                        if (context.ReceiveBuffer.Count >= 5 &&
            //                            (context.ReceiveBuffer[1] & 0x80) != 0)
            //                        {
            //                            // 错误响应帧长度为5字节
            //                            if (context.ReceiveBuffer.Count >= 5)
            //                            {
            //                                byte[] errorFrame = context.ReceiveBuffer.Take(5).ToArray();
            //                                ProcessCompleteFrame(errorFrame, context);
            //                                context.ReceiveBuffer.RemoveRange(0, 5);
            //                                continue;
            //                            }
            //                        }

            //                        // 检查正常响应帧
            //                        if (context.ReceiveBuffer[1] == 0x03) // 读取保持寄存器
            //                        {
            //                            // 正常响应帧长度 = 地址(1) + 功能码(1) + 字节数(1) + 数据(n*2) + CRC(2)
            //                            int expectedLength = 3 + context.ReceiveBuffer[2] + 2;

            //                            if (context.ReceiveBuffer.Count >= expectedLength)
            //                            {
            //                                byte[] completeFrame = context.ReceiveBuffer.Take(expectedLength).ToArray();
            //                                ProcessCompleteFrame(completeFrame, context);
            //                                context.ReceiveBuffer.RemoveRange(0, expectedLength);
            //                                continue;
            //                            }
            //                            else
            //                            {
            //                                // 帧不完整，等待更多数据
            //                                break;
            //                            }
            //                        }
            //                        else
            //                        {
            //                            // 未知帧类型，尝试查找下一个可能的帧头
            //                            int nextStart = FindNextFrameStart(context.ReceiveBuffer);
            //                            if (nextStart > 0)
            //                            {
            //                                // 移除无效数据
            //                                context.ReceiveBuffer.RemoveRange(0, nextStart);
            //                                continue;
            //                            }
            //                            else
            //                            {
            //                                // 没有找到有效帧头，清空缓冲区
            //                                context.ReceiveBuffer.Clear();
            //                                break;
            //                            }
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"ProcessReceivedData error: {ex.Message}");
            //                    context.ReceiveBuffer.Clear();
            //                }
            //            }

            //            private int FindNextFrameStart(List<byte> buffer)
            //            {
            //                for (int i = 1; i < buffer.Count; i++)
            //                {
            //                    // 有效的Modbus地址范围是1-247
            //                    if (buffer[i] >= 1 && buffer[i] <= 247)
            //                    {
            //                        // 检查是否是有效的功能码
            //                        if (i + 1 < buffer.Count)
            //                        {
            //                            byte functionCode = buffer[i + 1];
            //                            if (functionCode == 0x03 || (functionCode & 0x80) != 0)
            //                            {
            //                                return i;
            //                            }
            //                        }
            //                    }
            //                }
            //                return -1;
            //            }

            //            // 处理完整的帧
            //            private void ProcessCompleteFrame(byte[] frame, SerialPortContext context)
            //            {
            //                try
            //                {
            //                    // 显示接收到的原始数据
            //                    string rawDataHex = BitConverter.ToString(frame).Replace("-", " ");
            //                    string displayText = $"{context.Prefix}区（{context.Port.PortName}）收到{frame.Length}字节: {rawDataHex} -> {DateTime.Now:HH:mm:ss.fff} \r\n";

            //                    // 处理Modbus响应
            //                    bool isModbusResponse = ProcessModbusResponse(frame, context);

            //                    if (!isModbusResponse)
            //                    {
            //                        displayText += $"-> 非标准Modbus响应或CRC校验失败\r\n";
            //                    }

            //                    // 更新UI显示
            //                    if (!isPaused && !globalCts.IsCancellationRequested)
            //                    {
            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText(displayText);
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"ProcessCompleteFrame error: {ex.Message}");
            //                }
            //            }

            //            // 帧超时检测
            //            private void CheckFrameTimeout(SerialPortContext context)
            //            {
            //                lock (context.SerialLock)
            //                {
            //                    context.FrameTimer.Stop();

            //                    if (context.ReceiveBuffer.Count > 0)
            //                    {
            //                        // 帧超时，处理缓冲区中可能的不完整数据
            //                        string timeoutText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 帧超时，丢弃{context.ReceiveBuffer.Count}字节不完整数据\r\n";

            //                        if (!isPaused && !globalCts.IsCancellationRequested)
            //                        {
            //                            this.BeginInvoke(new Action(() =>
            //                            {
            //                                richTextBox1.AppendText(timeoutText);
            //                                richTextBox1.ScrollToCaret();
            //                            }));
            //                        }

            //                        context.ReceiveBuffer.Clear();
            //                    }

            //                    context.FrameTimer.Start();
            //                }
            //            }

            //            // 在 StartPollingAsync 方法中修改串口状态检查逻辑
            //            private async Task StartPollingAsync(SerialPortContext context)
            //            {
            //                int currentIndex = 0;
            //                int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "100");

            //                while (context.IsPolling && !context.Cts.Token.IsCancellationRequested && !globalCts.IsCancellationRequested)
            //                {
            //                    if (stationAddresses == null || stationAddresses.Length == 0)
            //                    {
            //                        await Task.Delay(100, context.Cts.Token);
            //                        continue;
            //                    }

            //                    int address = stationAddresses[currentIndex];
            //                    currentIndex = (currentIndex + 1) % stationAddresses.Length;

            //                    try
            //                    {
            //                        // 发送请求
            //                        byte[] request = BuildModbusRequest(address, context.StartRegisterAddress, context.RegisterCount, context);

            //                        // 获取该地址的等待器
            //                        var tcs = new TaskCompletionSource<bool>();
            //                        context.ResponseWaiters[address] = tcs;

            //                        using (var ctsTimeout = new CancellationTokenSource(RESPONSE_TIMEOUT))
            //                        {
            //                            var responseTask = tcs.Task;
            //                            var timeoutTask = Task.Delay(Timeout.Infinite, ctsTimeout.Token);

            //                            var completedTask = await Task.WhenAny(responseTask, timeoutTask);

            //                            if (completedTask == responseTask)
            //                            {
            //                                // 收到响应，取消超时
            //                                ctsTimeout.Cancel();
            //                                // 清除超时状态
            //                                portTimeoutStatus[context.Prefix] = false;
            //                                context.IsTimeout = false;
            //                            }
            //                            else
            //                            {
            //                                // 超时处理
            //                                context.ResponseWaiters.TryRemove(address, out _);
            //                                HandleResponseTimeout(context, address);
            //                            }
            //                        }
            //                    }
            //                    catch (OperationCanceledException)
            //                    {
            //                        // 取消操作，继续下一个
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"轮询错误: {ex.Message}");
            //                        await Task.Delay(1000, context.Cts.Token);
            //                    }

            //                    // 添加轮询间隔，避免过于频繁
            //                    await Task.Delay(pollingInterval, context.Cts.Token);
            //                }
            //            }

            //            // 处理响应超时
            //            private void HandleResponseTimeout(SerialPortContext context, int address)
            //            {
            //                if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                {
            //                    try
            //                    {
            //                        // 设置超时状态
            //                        portTimeoutStatus[context.Prefix] = true;
            //                        context.IsTimeout = true;

            //                        // 清空并设置报警消息
            //                        context.AlarmMessages.Clear();
            //                        string timeoutMessage = $"{context.Prefix}区({context.Port.PortName}) 报警器分线箱响应超时";
            //                        context.AlarmMessages.AppendLine(timeoutMessage);

            //                        // 更新报警消息标签 - 确保使用正确的标签名
            //                        string messageTag = $"meiqi{context.Prefix}";
            //                        hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());

            //                        // 更新所有寄存器的超时状态
            //                        for (int i = 1; i <= context.RegisterCount; i++)
            //                        {
            //                            hmi485.Tags[$"{context.Prefix}{i}_Value"].Write(999999); // 使用特殊值表示超时
            //                            hmi485.Tags[$"{context.Prefix}{i}_Time"].Write(DateTime.Now.ToString("离线时间：yyyy-MM-dd HH:mm:ss.fff"));
            //                        }

            //                        // 记录超时日志
            //                        Debug.WriteLine($"地址 {address} 响应超时，已更新WinCC标签");

            //                        if (!isPaused && !globalCts.IsCancellationRequested)
            //                        {
            //                            this.BeginInvoke(new Action(() =>
            //                            {
            //                                richTextBox1.AppendText($"{DateTime.Now:HH:mm:ss.fff} {context.Prefix}区{context.Port.PortName}: 报警器分线箱响应超时，已更新WinCC标签\r\n");
            //                                richTextBox1.ScrollToCaret();
            //                            }));
            //                        }
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"写入WinCC超时状态失败: {ex.Message}");

            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText($"写入WinCC超时状态失败: {ex.Message}\r\n");
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }
            //                }
            //            }

            //            private byte[] BuildModbusRequest(int address, int startRegister, int registerCount, SerialPortContext context)
            //            {
            //                byte[] request = new byte[8];
            //                request[0] = (byte)address;          // 设备地址
            //                request[1] = 0x03;                   // 功能码
            //                request[2] = (byte)(startRegister >> 8);  // 寄存器地址高字节
            //                request[3] = (byte)(startRegister & 0xFF);// 寄存器地址低字节
            //                request[4] = (byte)(registerCount >> 8);  // 读取数量高字节
            //                request[5] = (byte)(registerCount & 0xFF);// 读取数量低字节

            //                // 计算CRC校验
            //                byte[] crc;
            //                CRC_16(request.Take(6).ToArray(), out crc);
            //                request[6] = crc[0];
            //                request[7] = crc[1];

            //                // 将字节数组转换为十六进制字符串
            //                string hexString = BitConverter.ToString(request).Replace("-", " ");

            //                // 显示完整的请求帧
            //                if (!isPaused && !globalCts.IsCancellationRequested)
            //                {
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText($"{context.Prefix}区{context.Port.PortName}请求: {hexString}\r\n");
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }

            //                // 实际发送数据
            //                lock (context.SerialLock)
            //                {
            //                    if (context.Port.IsOpen && !globalCts.IsCancellationRequested)
            //                    {
            //                        try
            //                        {
            //                            context.Port.Write(request, 0, request.Length);
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"发送数据失败: {ex.Message}");
            //                        }
            //                    }
            //                }

            //                return request;
            //            }

            //            private bool ProcessModbusResponse(byte[] response, SerialPortContext context)
            //            {
            //                if (response.Length < 5) return false;

            //                try
            //                {
            //                    // 在处理响应前清空报警消息
            //                    context.AlarmMessages.Clear();

            //                    // 验证CRC
            //                    byte[] crc;
            //                    CRC_16(response.Take(response.Length - 2).ToArray(), out crc);
            //                    if (crc[0] != response[response.Length - 2] || crc[1] != response[response.Length - 1])
            //                    {
            //                        Debug.WriteLine("CRC校验失败");
            //                        return false;
            //                    }

            //                    int address = response[0]; // 设备地址
            //                    int functionCode = response[1];

            //                    // 检查错误响应
            //                    if ((functionCode & 0x80) != 0)
            //                    {
            //                        int errorCode = response[2];
            //                        string errorMessage = GetModbusErrorDescription(errorCode);
            //                        string errorText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 地址{address}错误响应: {errorMessage}\r\n";

            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText(errorText);
            //                            richTextBox1.ScrollToCaret();
            //                        }));

            //                        return true;
            //                    }

            //                    if (functionCode == 0x03) // 读取保持寄存器响应
            //                    {
            //                        int byteCount = response[2];
            //                        if (byteCount != context.RegisterCount * 2) return false;

            //                        // 通知等待器
            //                        if (context.ResponseWaiters.TryRemove(address, out var tcs))
            //                        {
            //                            tcs.TrySetResult(true);
            //                        }

            //                        // 清除超时状态
            //                        portTimeoutStatus[context.Prefix] = false;
            //                        context.IsTimeout = false;

            //                        // 解析所有寄存器值并更新到WinCC
            //                        for (int i = 0; i < context.RegisterCount; i++)
            //                        {
            //                            int registerIndex = 3 + i * 2; // 数据起始位置
            //                            int value = (response[registerIndex] << 8) | response[registerIndex + 1];
            //                            int registerNumber = i + 1; // 寄存器编号从1开始

            //                            // 处理特殊值和更新WinCC
            //                            ProcessRegisterValue(address, registerNumber, value, context);
            //                        }

            //                        // 在处理完所有寄存器后更新报警消息
            //                        if (context.AlarmMessages.Length > 0)
            //                        {
            //                            try
            //                            {
            //                                string messageTag = $"meiqi{context.Prefix}";
            //                                hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());
            //                                context.LastAlarmUpdate = DateTime.Now;
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                Debug.WriteLine($"更新报警消息失败: {ex.Message}");
            //                            }
            //                        }
            //                        else
            //                        {
            //                            // 如果没有报警，清空报警消息
            //                            try
            //                            {
            //                                string messageTag = $"meiqi{context.Prefix}";
            //                                hmi485.Tags[messageTag].Write("无报警");
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                Debug.WriteLine($"清空报警消息失败: {ex.Message}");
            //                            }
            //                        }

            //                        return true;
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}区响应，报警消息长度: {context.AlarmMessages.Length}");
            //                    Debug.WriteLine($"ProcessModbusResponse error: {ex.Message}");
            //                }

            //                return false;
            //            }

            //            private void ProcessRegisterValue(int address, int registerNumber, int value, SerialPortContext context)
            //            {
            //                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //                string displayText = $"{context.Prefix}区（{context.Port.PortName}）第{registerNumber}路数据: 值：{value} 读取时间：{timestamp}，";
            //                string messageTag = $"meiqi{context.Prefix}";
            //                try
            //                {
            //                    if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                    {
            //                        string statusTag = $"meiqi{context.Prefix}区";
            //                        string valueTag = $"{context.Prefix}{registerNumber}_Value";
            //                        string timeTag = $"{context.Prefix}{registerNumber}_Time";

            //                        // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                        if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                        {
            //                            displayText += $"-> 串口超时，不更新WinCC状态\r\n";
            //                            return;
            //                        }

            //                        hmi485.Tags[statusTag].Write(1);

            //                        // 处理特殊值
            //                        if (value == 0xFAFA) // 探测器没有接
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFAFA);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 没有连接");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器没有连接；");
            //                            displayText += $"-> 探测器没有连接\r\n";
            //                        }
            //                        else if (value == 0xFEFE) // 探测器丢失
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFEFE);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 丢失");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器丢失");
            //                            displayText += $"-> 探测器丢失\r\n";
            //                        }
            //                        else if (value == 0xFCFC) // 探测器故障
            //                        {
            //                            hmi485.Tags[valueTag].Write(0xFCFC);
            //                            hmi485.Tags[timeTag].Write($"{timestamp} - 故障");
            //                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器故障");
            //                            displayText += $"-> 探测器故障\r\n";
            //                        }
            //                        else // 正常值
            //                        {
            //                            float actualValue = value / 10.0f;
            //                            hmi485.Tags[valueTag].Write(actualValue);
            //                            hmi485.Tags[timeTag].Write(timestamp);
            //                            hmi485.Tags[messageTag].Write($"CO报警地址 {context.Prefix}{registerNumber}  值: {actualValue:F1}ppm");
            //                            // 正常值不添加到报警消息中
            //                            string tagName = $"meiqi{context.Prefix}区";
            //                            hmi485.Tags[tagName].Write(1);
            //                        }
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}{registerNumber}，值: 0x{value:X4}");
            //                    displayText += $"! 写入WinCC标签失败: {ex.Message}\r\n";
            //                }

            //                // 更新UI
            //                if (!isPaused && !globalCts.IsCancellationRequested)
            //                {
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText(displayText);
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }

            //            private string GetModbusErrorDescription(int errorCode)
            //            {
            //                switch (errorCode)
            //                {
            //                    case 0x01: return "非法功能";
            //                    case 0x02: return "非法数据地址";
            //                    case 0x03: return "非法数据值";
            //                    case 0x04: return "从站设备故障";
            //                    case 0x05: return "确认";
            //                    case 0x06: return "从属设备忙";
            //                    case 0x08: return "存储奇偶性差错";
            //                    case 0x0A: return "不可用网关路径";
            //                    case 0x0B: return "网关目标设备响应失败";
            //                    default: return "未知错误";
            //                }
            //            }

            //            public static byte[] CRC_16(byte[] data, out byte[] temdata)
            //            {
            //                if (data.Length == 0)
            //                    throw new Exception("调用CRC16校验算法,（低字节在前，高字节在后）时发生异常，异常信息：被校验的数组长度为0。");
            //                int xda, xdapoly;
            //                byte i, j, xdabit;
            //                xda = 0xFFFF;
            //                xdapoly = 0xA001;
            //                for (i = 0; i < data.Length; i++)
            //                {
            //                    xda ^= data[i];
            //                    for (j = 0; j < 8; j++)
            //                    {
            //                        xdabit = (byte)(xda & 0x01);
            //                        xda >>= 1;
            //                        if (xdabit == 1)
            //                            xda ^= xdapoly;
            //                    }
            //                }
            //                temdata = new byte[2] { (byte)(xda & 0xFF), (byte)(xda >> 8) };
            //                return temdata;
            //            }

            //            private void Form1_FormClosing(object sender, FormClosingEventArgs e)
            //            {
            //                // 停止所有后台操作
            //                globalCts.Cancel();
            //                twoTimer.Stop();

            //                // 停止所有串口轮询并关闭串口
            //                foreach (var context in serialPortContexts)
            //                {
            //                    context.IsPolling = false;
            //                    context.Cts.Cancel();
            //                    context.FrameTimer.Stop();
            //                    if (context.Port.IsOpen) context.Port.Close();
            //                    context.Dispose();
            //                }

            //                // 更新WinCC状态为关闭 - 确保所有前缀的状态都被更新
            //                try
            //                {
            //                    if (!stopWritingToWinCC)
            //                    {
            //                        foreach (var context in serialPortContexts)
            //                        {
            //                            try
            //                            {
            //                                //string tagName = $"meiqi{context.Prefix}区";
            //                                //hmi485.Tags[tagName].Write(0);
            //                            }
            //                            catch { /* 忽略单个串口关闭时的异常 */ }
            //                        }
            //                    }
            //                }
            //                catch { /* 忽略关闭时的异常 */ }
            //            }

            //            private void richTextBox1_TextChanged(object sender, EventArgs e)
            //            {
            //                if (isPaused || globalCts.IsCancellationRequested) return;

            //                const int maxLine = 1200;
            //                if (richTextBox1.Lines.Length > maxLine)
            //                {
            //                    int firstNewLine = richTextBox1.Text.IndexOf('\n');
            //                    if (firstNewLine >= 0)
            //                    {
            //                        // 保留滚动位置
            //                        int scrollPos = richTextBox1.GetCharIndexFromPosition(new Point(0, 0));
            //                        richTextBox1.Text = richTextBox1.Text.Substring(firstNewLine + 1);
            //                        richTextBox1.ScrollToCaret();
            //                        if (scrollPos > 0)
            //                        {
            //                            richTextBox1.SelectionStart = Math.Min(scrollPos - firstNewLine - 1, richTextBox1.Text.Length);
            //                            richTextBox1.ScrollToCaret();
            //                        }
            //                    }
            //                }
            //            }

            //            protected override void Dispose(bool disposing)
            //            {
            //                if (disposing)
            //                {
            //                    try
            //                    {
            //                        // 取消所有操作
            //                        globalCts?.Cancel();

            //                        // 停止并释放定时器
            //                        twoTimer?.Stop();
            //                        twoTimer?.Dispose();

            //                        // 释放所有串口上下文
            //                        foreach (var context in serialPortContexts)
            //                        {
            //                            context?.Dispose();
            //                        }
            //                        serialPortContexts.Clear();

            //                        // 释放组件资源
            //                        components?.Dispose();
            //                    }
            //                    catch
            //                    {
            //                        // 忽略清理过程中的异常
            //                    }
            //                    finally
            //                    {
            //                        globalCts?.Dispose();
            //                    }
            //                }
            //                base.Dispose(disposing);
            //            }

            //            private void button1_Click(object sender, EventArgs e) { isPaused = !isPaused; button1.Text = isPaused ? "继续滚屏显示" : "停止滚屏显示"; }
            //            private void textBox1_TextChanged(object sender, EventArgs e) { }
            //        }
            //    }
            #endregion//不再更新掉线的通讯状态
            #region//0.4
            //using CCHMIRUNTIME;
            //using System;
            //using System.Collections.Concurrent;
            //using System.Collections.Generic;
            //using System.Configuration;
            //using System.Data;
            //using System.Diagnostics;
            //using System.Drawing;
            //using System.IO.Ports;
            //using System.Linq;
            //using System.Text;
            //using System.Threading;
            //using System.Threading.Tasks;
            //using System.Windows.Forms;

            //namespace meiqibaojing
            //{
            //    public partial class Form1 : Form
            //    {
            //        int RESPONSE_TIMEOUT = int.Parse(ConfigurationManager.AppSettings["RESPONSETIMEOUT"]);
            //        private int[] stationAddresses; // 站地址数组
            //        private HMIRuntime hmi485 = new CCHMIRUNTIME.HMIRuntime();
            //        private bool isPaused = false;
            //        private bool stopWritingToWinCC = false;
            //        private CancellationTokenSource globalCts = new CancellationTokenSource();
            //        private List<SerialPortContext> serialPortContexts = new List<SerialPortContext>();
            //        private System.Windows.Forms.Timer twoTimer = new System.Windows.Forms.Timer();
            //        // 添加超时状态字典
            //        private ConcurrentDictionary<string, bool> portTimeoutStatus = new ConcurrentDictionary<string, bool>();

            //        private class SerialPortContext : IDisposable
            //        {
            //            public StringBuilder AlarmMessages { get; } = new StringBuilder();
            //            public DateTime LastAlarmUpdate { get; set; } = DateTime.MinValue;
            //            public SerialPort Port { get; set; }
            //            public string Prefix { get; set; }
            //            public int StartRegisterAddress { get; set; }
            //            public int RegisterCount { get; set; }
            //            public List<byte> ReceiveBuffer { get; } = new List<byte>();
            //            public DateTime LastReceivedTime { get; set; } = DateTime.MinValue;
            //            public ConcurrentDictionary<int, TaskCompletionSource<bool>> ResponseWaiters { get; } = new ConcurrentDictionary<int, TaskCompletionSource<bool>>();
            //            public object SerialLock { get; } = new object();
            //            public bool IsPolling { get; set; } = true;
            //            public CancellationTokenSource Cts { get; } = new CancellationTokenSource();
            //            public int Index { get; set; }  // 串口索引
            //            public System.Windows.Forms.Timer FrameTimer { get; set; } // 帧超时定时器
            //            public bool IsTimeout { get; set; } = false; // 添加超时状态

            //            public void Dispose()
            //            {
            //                try
            //                {
            //                    Cts?.Cancel();
            //                    Cts?.Dispose();
            //                    FrameTimer?.Stop();
            //                    FrameTimer?.Dispose();

            //                    // 确保串口完全释放
            //                    if (Port != null)
            //                    {
            //                        try
            //                        {
            //                            if (Port.IsOpen)
            //                                Port.Close();
            //                            Port.Dispose();
            //                            Port = null; // 设置为null防止重复使用
            //                        }
            //                        catch
            //                        {
            //                            // 忽略释放异常
            //                            Port = null;
            //                        }
            //                    }
            //                }
            //                catch
            //                {
            //                    // 忽略所有异常
            //                }
            //            }
            //        }// 串口上下文类，管理每个串口的独立状态

            //        public Form1()
            //        {
            //            InitializeComponent();
            //            LoadPorts();
            //            twoTimer.Interval = 1000;
            //            twoTimer.Tick += new EventHandler(WinccConn);
            //            twoTimer.Start();
            //        }

            //        private async void Form1_Load(object sender, EventArgs e)
            //        {
            //            Task while04 = Task.Run(() => Winccstart());
            //            try
            //            {
            //                // 初始化所有串口
            //                foreach (var context in serialPortContexts)
            //                {
            //                    try
            //                    {
            //                        context.Port.DataReceived += (s, args) => SerialPort_DataReceived(s, args, context);
            //                        context.Port.Open();
            //                        context.FrameTimer.Start(); // 启动帧超时检测定时器

            //                        // 更新WinCC状态 - 确保每个串口的状态都被更新
            //                        string tagName = $"meiqi{context.Prefix}区";
            //                        hmi485.Tags[tagName].Write(1);
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        MessageBox.Show($"初始化串口{context.Port.PortName}失败: {ex.Message}");
            //                        continue;
            //                    }
            //                }

            //                // 启动轮询任务
            //                var pollingTasks = serialPortContexts
            //                    .Where(context => context.Port.IsOpen)
            //                    .Select(context => StartPollingAsync(context));

            //                await Task.WhenAll(pollingTasks);
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show($"初始化失败: {ex.Message}");
            //                Environment.Exit(Environment.ExitCode);
            //            }
            //        }

            //        private void LoadPorts()
            //        {
            //            try
            //            {
            //                int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "200");
            //                string addressesConfig = ConfigurationManager.AppSettings["StationAddresses"];
            //                if (!string.IsNullOrEmpty(addressesConfig))
            //                {
            //                    stationAddresses = addressesConfig.Split(',')
            //                        .Select(addr => int.Parse(addr.Trim()))
            //                        .ToArray();
            //                }
            //                else
            //                {
            //                    stationAddresses = new int[] { 1 };
            //                }

            //                string[] portNames = ConfigurationManager.AppSettings["SerialPorts"].Split(';');
            //                int[] baudRates = ConfigurationManager.AppSettings["BaudRates"].Split(';').Select(int.Parse).ToArray();
            //                Parity[] parities = ConfigurationManager.AppSettings["Parities"].Split(';').Select(p => (Parity)Enum.Parse(typeof(Parity), p)).ToArray();
            //                int[] dataBitsArray = ConfigurationManager.AppSettings["DataBitsArray"].Split(';').Select(int.Parse).ToArray();
            //                StopBits[] stopBitsArray = ConfigurationManager.AppSettings["StopBitsArray"].Split(';').Select(s => (StopBits)Enum.Parse(typeof(StopBits), s)).ToArray();
            //                string[] prefixes = ConfigurationManager.AppSettings["AddressPrefixes"].Split(';');
            //                int[] startRegisterAddresses = ConfigurationManager.AppSettings["RegisterAddresses"].Split(';').Select(int.Parse).ToArray();
            //                int[] registerCounts;
            //                string registerCountsConfig = ConfigurationManager.AppSettings["RegisterCounts"];
            //                if (!string.IsNullOrEmpty(registerCountsConfig))
            //                {
            //                    registerCounts = registerCountsConfig.Split(';').Select(int.Parse).ToArray();
            //                }
            //                else
            //                {
            //                    registerCounts = new int[portNames.Length];
            //                    for (int i = 0; i < portNames.Length; i++)
            //                    {
            //                        registerCounts[i] = 16;
            //                    }
            //                }
            //                if (new[] { portNames.Length, baudRates.Length, parities.Length, dataBitsArray.Length, stopBitsArray.Length, prefixes.Length, startRegisterAddresses.Length, registerCounts.Length }
            //                   .Distinct().Count() != 1)
            //                {
            //                    throw new ConfigurationErrorsException("串口配置数组长度不一致");
            //                }

            //                StringBuilder configText = new StringBuilder("串口配置:\r\n");
            //                for (int i = 0; i < portNames.Length; i++)
            //                {
            //                    configText.AppendLine($"串口{i + 1}: {portNames[i]}, {baudRates[i]}, {parities[i]}, {dataBitsArray[i]}, {stopBitsArray[i]}, 区域:{prefixes[i]}, 起始寄存器:{startRegisterAddresses[i]}, 读取数量:{registerCounts[i]}, 轮询间隔:{pollingInterval}ms");// 创建串口上下文
            //                    var context = new SerialPortContext
            //                    {
            //                        Port = new SerialPort(portNames[i], baudRates[i], parities[i], dataBitsArray[i], stopBitsArray[i]),
            //                        Prefix = prefixes[i],
            //                        StartRegisterAddress = startRegisterAddresses[i],
            //                        RegisterCount = registerCounts[i],
            //                        Index = i + 1,
            //                        FrameTimer = new System.Windows.Forms.Timer { Interval = int.Parse(ConfigurationManager.AppSettings["FrameInterval"] ?? "20") }
            //                    };

            //                    context.FrameTimer.Tick += (sender, e) => CheckFrameTimeout(context);
            //                    serialPortContexts.Add(context);
            //                    // 初始化超时状态
            //                    portTimeoutStatus[context.Prefix] = false;
            //                }

            //                textBox1.Text = configText.ToString();
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show($"加载配置失败: {ex.Message}");
            //                Environment.Exit(1);
            //            }
            //        }

            //        private void WinccConn(object sender, EventArgs e)
            //        {
            //            try
            //            {
            //                if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;

            //                foreach (var context in serialPortContexts)
            //                {
            //                    try
            //                    {
            //                        // 添加 null 检查
            //                        if (context == null || context.Port == null)
            //                        {
            //                            continue;
            //                        }

            //                        // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                        if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                        {
            //                            continue;
            //                        }

            //                        if (context.Port.IsOpen)
            //                        {
            //                            string tagName = $"meiqi{context.Prefix}区";
            //                            hmi485.Tags[tagName]?.Write(1);
            //                        }
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"更新串口状态到WinCC失败: {ex.Message}");
            //                        // 即使出错也继续处理其他串口
            //                    }
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;
            //                MessageBox.Show($"与WinCC通讯失败: {ex.Message}，驱动程序将退出。");
            //                Environment.Exit(Environment.ExitCode);
            //            }
            //        }

            //        private async Task Winccstart()
            //        {
            //            string processName = ConfigurationManager.AppSettings["ProcessName"];

            //            while (!globalCts.IsCancellationRequested)
            //            {
            //                try
            //                {
            //                    bool b = IsProcessStarted(processName);

            //                    if (b)
            //                    {
            //                        stopWritingToWinCC = false;
            //                        // 更新所有串口的状态
            //                        foreach (var context in serialPortContexts)
            //                        {
            //                            try
            //                            {
            //                                // 添加 null 检查
            //                                if (context == null || context.Port == null)
            //                                {
            //                                    continue;
            //                                }

            //                                // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                                if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                                {
            //                                    continue;
            //                                }

            //                                if (context.Port.IsOpen)
            //                                {
            //                                    string tagName = $"meiqi{context.Prefix}区";
            //                                    hmi485.Tags[tagName]?.Write(1);
            //                                }
            //                            }
            //                            catch (Exception ex)
            //                            {
            //                                Debug.WriteLine($"更新串口{context?.Port?.PortName}状态到WinCC失败: {ex.Message}");
            //                            }
            //                        }
            //                    }
            //                    else
            //                    {
            //                        if (InvokeRequired)
            //                        {
            //                            Invoke(new Action(() =>
            //                            {
            //                                DialogResult result = MessageBox.Show(
            //                                    $"未找到运行的WinCC实例({processName})，是否退出驱动程序？",
            //                                    "WinCC未运行",
            //                                    MessageBoxButtons.YesNo,
            //                                    MessageBoxIcon.Warning);

            //                                if (result == DialogResult.Yes)
            //                                {
            //                                    MessageBox.Show("驱动程序将退出。");
            //                                    Environment.Exit(Environment.ExitCode);
            //                                }
            //                                else
            //                                {
            //                                    stopWritingToWinCC = true;
            //                                    richTextBox1.AppendText($"已停止向WinCC写入标签。WinCC进程({processName})未运行。\r\n");
            //                                    richTextBox1.ScrollToCaret();
            //                                }
            //                            }));
            //                        }

            //                        await Task.Delay(10000, globalCts.Token);
            //                    }

            //                    await Task.Delay(1000, globalCts.Token);
            //                }
            //                catch (OperationCanceledException)
            //                {
            //                    // 任务被取消，正常退出
            //                    return;
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"WinCC进程检测错误: {ex.Message}");
            //                    await Task.Delay(5000, globalCts.Token);
            //                }
            //            }
            //        }

            //        bool IsProcessStarted(string processName1)
            //        {
            //            try
            //            {
            //                Process[] temp = Process.GetProcessesByName(processName1);
            //                return temp.Length > 0;
            //            }
            //            catch
            //            {
            //                return false;
            //            }
            //        }

            //        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e, SerialPortContext context)
            //        {
            //            try
            //            {
            //                lock (context.SerialLock)
            //                {
            //                    int bytesToRead = context.Port.BytesToRead;
            //                    if (bytesToRead <= 0) return;

            //                    byte[] buffer = new byte[bytesToRead];
            //                    context.Port.Read(buffer, 0, bytesToRead);

            //                    // 更新最后接收时间
            //                    context.LastReceivedTime = DateTime.Now;

            //                    // 将数据添加到缓冲区
            //                    context.ReceiveBuffer.AddRange(buffer);

            //                    // 重置帧超时定时器
            //                    context.FrameTimer.Stop();
            //                    context.FrameTimer.Start();

            //                    // 尝试处理缓冲区中的完整帧
            //                    ProcessReceivedData(context);
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"SerialPort_DataReceived error: {ex.Message}");
            //                this.BeginInvoke(new Action(() =>
            //                {
            //                    richTextBox1.AppendText($"串口{context.Port.PortName}接收数据错误: {ex.Message}\r\n");
            //                    richTextBox1.ScrollToCaret();
            //                }));
            //            }
            //        }

            //        private void ProcessReceivedData(SerialPortContext context)
            //        {
            //            try
            //            {
            //                if (context.ReceiveBuffer.Count == 0) return;

            //                // 查找可能的完整帧
            //                while (context.ReceiveBuffer.Count >= 5) // 最小帧长度
            //                {
            //                    // 检查是否为错误响应帧（5字节）
            //                    if (context.ReceiveBuffer.Count >= 5 &&
            //                        (context.ReceiveBuffer[1] & 0x80) != 0)
            //                    {
            //                        // 错误响应帧长度为5字节
            //                        if (context.ReceiveBuffer.Count >= 5)
            //                        {
            //                            byte[] errorFrame = context.ReceiveBuffer.Take(5).ToArray();
            //                            ProcessCompleteFrame(errorFrame, context);
            //                            context.ReceiveBuffer.RemoveRange(0, 5);
            //                            continue;
            //                        }
            //                    }

            //                    // 检查正常响应帧
            //                    if (context.ReceiveBuffer[1] == 0x03) // 读取保持寄存器
            //                    {
            //                        // 正常响应帧长度 = 地址(1) + 功能码(1) + 字节数(1) + 数据(n*2) + CRC(2)
            //                        int expectedLength = 3 + context.ReceiveBuffer[2] + 2;

            //                        if (context.ReceiveBuffer.Count >= expectedLength)
            //                        {
            //                            byte[] completeFrame = context.ReceiveBuffer.Take(expectedLength).ToArray();
            //                            ProcessCompleteFrame(completeFrame, context);
            //                            context.ReceiveBuffer.RemoveRange(0, expectedLength);
            //                            continue;
            //                        }
            //                        else
            //                        {
            //                            // 帧不完整，等待更多数据
            //                            break;
            //                        }
            //                    }
            //                    else
            //                    {
            //                        // 未知帧类型，尝试查找下一个可能的帧头
            //                        int nextStart = FindNextFrameStart(context.ReceiveBuffer);
            //                        if (nextStart > 0)
            //                        {
            //                            // 移除无效数据
            //                            context.ReceiveBuffer.RemoveRange(0, nextStart);
            //                            continue;
            //                        }
            //                        else
            //                        {
            //                            // 没有找到有效帧头，清空缓冲区
            //                            context.ReceiveBuffer.Clear();
            //                            break;
            //                        }
            //                    }
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"ProcessReceivedData error: {ex.Message}");
            //                context.ReceiveBuffer.Clear();
            //            }
            //        }

            //        private int FindNextFrameStart(List<byte> buffer)
            //        {
            //            for (int i = 1; i < buffer.Count; i++)
            //            {
            //                // 有效的Modbus地址范围是1-247
            //                if (buffer[i] >= 1 && buffer[i] <= 247)
            //                {
            //                    // 检查是否是有效的功能码
            //                    if (i + 1 < buffer.Count)
            //                    {
            //                        byte functionCode = buffer[i + 1];
            //                        if (functionCode == 0x03 || (functionCode & 0x80) != 0)
            //                        {
            //                            return i;
            //                        }
            //                    }
            //                }
            //            }
            //            return -1;
            //        }

            //        // 处理完整的帧
            //        // 处理完整的帧
            //        private void ProcessCompleteFrame(byte[] frame, SerialPortContext context)
            //        {
            //            try
            //            {
            //                // 显示接收到的原始数据
            //                string rawDataHex = BitConverter.ToString(frame).Replace("-", " ");
            //                string timestamp = DateTime.Now.ToString("mm:ss");
            //                string displayText = $"{context.Prefix}区（{context.Port.PortName}）收到{frame.Length}字节: {rawDataHex} -> {timestamp} \r\n";

            //                // 写入WinCC接收报文变量
            //                try
            //                {
            //                    if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                    {
            //                        string messageTag = $"meiqi{context.Prefix}1";
            //                        string receiveMessage = $"{timestamp}=>{rawDataHex}";
            //                        hmi485.Tags[messageTag].Write(receiveMessage);
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"写入WinCC接收报文失败: {ex.Message}");
            //                }

            //                // 处理Modbus响应
            //                bool isModbusResponse = ProcessModbusResponse(frame, context);

            //                if (!isModbusResponse)
            //                {
            //                    displayText += $"-> 非标准Modbus响应或CRC校验失败\r\n";
            //                }

            //                // 更新UI显示
            //                if (!isPaused && !globalCts.IsCancellationRequested)
            //                {
            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText(displayText);
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"ProcessCompleteFrame error: {ex.Message}");
            //            }
            //        }

            //        // 帧超时检测
            //        private void CheckFrameTimeout(SerialPortContext context)
            //        {
            //            lock (context.SerialLock)
            //            {
            //                context.FrameTimer.Stop();

            //                if (context.ReceiveBuffer.Count > 0)
            //                {
            //                    // 帧超时，处理缓冲区中可能的不完整数据
            //                    string timeoutText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 帧超时，丢弃{context.ReceiveBuffer.Count}字节不完整数据\r\n";

            //                    if (!isPaused && !globalCts.IsCancellationRequested)
            //                    {
            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText(timeoutText);
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }

            //                    context.ReceiveBuffer.Clear();
            //                }

            //                context.FrameTimer.Start();
            //            }
            //        }

            //        // 在 StartPollingAsync 方法中修改串口状态检查逻辑
            //        private async Task StartPollingAsync(SerialPortContext context)
            //        {
            //            int currentIndex = 0;
            //            int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "100");

            //            while (context.IsPolling && !context.Cts.Token.IsCancellationRequested && !globalCts.IsCancellationRequested)
            //            {
            //                if (stationAddresses == null || stationAddresses.Length == 0)
            //                {
            //                    await Task.Delay(100, context.Cts.Token);
            //                    continue;
            //                }

            //                int address = stationAddresses[currentIndex];
            //                currentIndex = (currentIndex + 1) % stationAddresses.Length;

            //                try
            //                {
            //                    // 发送请求
            //                    byte[] request = BuildModbusRequest(address, context.StartRegisterAddress, context.RegisterCount, context);

            //                    // 获取该地址的等待器
            //                    var tcs = new TaskCompletionSource<bool>();
            //                    context.ResponseWaiters[address] = tcs;

            //                    using (var ctsTimeout = new CancellationTokenSource(RESPONSE_TIMEOUT))
            //                    {
            //                        var responseTask = tcs.Task;
            //                        var timeoutTask = Task.Delay(Timeout.Infinite, ctsTimeout.Token);

            //                        var completedTask = await Task.WhenAny(responseTask, timeoutTask);

            //                        if (completedTask == responseTask)
            //                        {
            //                            // 收到响应，取消超时
            //                            ctsTimeout.Cancel();
            //                            // 清除超时状态
            //                            portTimeoutStatus[context.Prefix] = false;
            //                            context.IsTimeout = false;
            //                        }
            //                        else
            //                        {
            //                            // 超时处理
            //                            context.ResponseWaiters.TryRemove(address, out _);
            //                            HandleResponseTimeout(context, address);
            //                        }
            //                    }
            //                }
            //                catch (OperationCanceledException)
            //                {
            //                    // 取消操作，继续下一个
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"轮询错误: {ex.Message}");
            //                    await Task.Delay(1000, context.Cts.Token);
            //                }

            //                // 添加轮询间隔，避免过于频繁
            //                await Task.Delay(pollingInterval, context.Cts.Token);
            //            }
            //        }

            //        // 处理响应超时
            //        // 处理响应超时
            //        private void HandleResponseTimeout(SerialPortContext context, int address)
            //        {
            //            if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //            {
            //                try
            //                {
            //                    // 清空并设置报警消息
            //                    context.AlarmMessages.Clear();
            //                    string timeoutMessage = $"{context.Prefix}区({context.Port.PortName}) 报警器分线箱离线";
            //                    context.AlarmMessages.AppendLine(timeoutMessage);

            //                    // 写入超时信息到WinCC报文变量
            //                    string timestamp = DateTime.Now.ToString("mm:ss");
            //                    string timeoutInfo = $"{timestamp}=>{context.Prefix}区 串口{context.Port.PortName}无响应";
            //                    string messageTag = $"meiqi{context.Prefix}1";
            //                    hmi485.Tags[messageTag].Write(timeoutInfo);

            //                    // 更新报警消息标签
            //                    hmi485.Tags[$"meiqi{context.Prefix}"].Write(context.AlarmMessages.ToString());

            //                    // 更新所有寄存器的超时状态
            //                    for (int i = 1; i <= context.RegisterCount; i++)
            //                    {
            //                        hmi485.Tags[$"{context.Prefix}{i}_Value"].Write(999999); // 使用特殊值表示超时
            //                        hmi485.Tags[$"{context.Prefix}{i}_Time"].Write(DateTime.Now.ToString("离线时间：yyyy-MM-dd HH:mm:ss.fff"));
            //                    }

            //                    // 记录超时日志
            //                    Debug.WriteLine($"地址 {address} 响应超时，已更新WinCC标签");

            //                    if (!isPaused && !globalCts.IsCancellationRequested)
            //                    {
            //                        this.BeginInvoke(new Action(() =>
            //                        {
            //                            richTextBox1.AppendText($"{timestamp} {context.Prefix}区{context.Port.PortName}: 报警器分线箱响应超时，已更新WinCC标签\r\n");
            //                            richTextBox1.ScrollToCaret();
            //                        }));
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Debug.WriteLine($"写入WinCC超时状态失败: {ex.Message}");

            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText($"写入WinCC超时状态失败: {ex.Message}\r\n");
            //                        richTextBox1.ScrollToCaret();
            //                    }));
            //                }
            //            }
            //        }

            //        private byte[] BuildModbusRequest(int address, int startRegister, int registerCount, SerialPortContext context)
            //        {
            //            byte[] request = new byte[8];
            //            request[0] = (byte)address;          // 设备地址
            //            request[1] = 0x03;                   // 功能码
            //            request[2] = (byte)(startRegister >> 8);  // 寄存器地址高字节
            //            request[3] = (byte)(startRegister & 0xFF);// 寄存器地址低字节
            //            request[4] = (byte)(registerCount >> 8);  // 读取数量高字节
            //            request[5] = (byte)(registerCount & 0xFF);// 读取数量低字节

            //            // 计算CRC校验
            //            byte[] crc;
            //            CRC_16(request.Take(6).ToArray(), out crc);
            //            request[6] = crc[0];
            //            request[7] = crc[1];

            //            // 将字节数组转换为十六进制字符串
            //            string hexString = BitConverter.ToString(request).Replace("-", " ");
            //            string timestamp = DateTime.Now.ToString("mm:ss");
            //            string requestMessage = $"{timestamp}=>{hexString}";

            //            // 写入WinCC变量
            //            try
            //            {
            //                if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                {
            //                    string messageTag = $"meiqi{context.Prefix}1";
            //                    hmi485.Tags[messageTag].Write(requestMessage);
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"写入WinCC发送报文失败: {ex.Message}");
            //            }

            //            // 显示完整的请求帧
            //            if (!isPaused && !globalCts.IsCancellationRequested)
            //            {
            //                this.BeginInvoke(new Action(() =>
            //                {
            //                    richTextBox1.AppendText($"{context.Prefix}区{context.Port.PortName}请求: {hexString}\r\n");
            //                    richTextBox1.ScrollToCaret();
            //                }));
            //            }

            //            // 实际发送数据
            //            lock (context.SerialLock)
            //            {
            //                if (context.Port.IsOpen && !globalCts.IsCancellationRequested)
            //                {
            //                    try
            //                    {
            //                        context.Port.Write(request, 0, request.Length);
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"发送数据失败: {ex.Message}");
            //                    }
            //                }
            //            }

            //            return request;
            //        }

            //        private bool ProcessModbusResponse(byte[] response, SerialPortContext context)
            //        {
            //            if (response.Length < 5) return false;

            //            try
            //            {
            //                // 在处理响应前清空报警消息
            //                context.AlarmMessages.Clear();

            //                // 验证CRC
            //                byte[] crc;
            //                CRC_16(response.Take(response.Length - 2).ToArray(), out crc);
            //                if (crc[0] != response[response.Length - 2] || crc[1] != response[response.Length - 1])
            //                {
            //                    Debug.WriteLine("CRC校验失败");
            //                    return false;
            //                }

            //                int address = response[0]; // 设备地址
            //                int functionCode = response[1];

            //                // 检查错误响应
            //                if ((functionCode & 0x80) != 0)
            //                {
            //                    int errorCode = response[2];
            //                    string errorMessage = GetModbusErrorDescription(errorCode);
            //                    string errorText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 地址{address}错误响应: {errorMessage}\r\n";

            //                    // 写入错误信息到WinCC
            //                    try
            //                    {
            //                        if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                        {
            //                            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            //                            string errorInfo = $"[错误] {timestamp} 地址{address}: {errorMessage}";
            //                            string messageTag = $"meiqi{context.Prefix}1";
            //                            hmi485.Tags[messageTag].Write(errorInfo);
            //                        }
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        Debug.WriteLine($"写入WinCC错误信息失败: {ex.Message}");
            //                    }

            //                    this.BeginInvoke(new Action(() =>
            //                    {
            //                        richTextBox1.AppendText(errorText);
            //                        richTextBox1.ScrollToCaret();
            //                    }));

            //                    return true;
            //                }

            //                if (functionCode == 0x03) // 读取保持寄存器响应
            //                {
            //                    int byteCount = response[2];
            //                    if (byteCount != context.RegisterCount * 2) return false;

            //                    // 通知等待器
            //                    if (context.ResponseWaiters.TryRemove(address, out var tcs))
            //                    {
            //                        tcs.TrySetResult(true);
            //                    }

            //                    // 清除超时状态
            //                    portTimeoutStatus[context.Prefix] = false;
            //                    context.IsTimeout = false;

            //                    // 解析所有寄存器值并更新到WinCC
            //                    for (int i = 0; i < context.RegisterCount; i++)
            //                    {
            //                        int registerIndex = 3 + i * 2; // 数据起始位置
            //                        int value = (response[registerIndex] << 8) | response[registerIndex + 1];
            //                        int registerNumber = i + 1; // 寄存器编号从1开始

            //                        // 处理特殊值和更新WinCC
            //                        ProcessRegisterValue(address, registerNumber, value, context);
            //                    }

            //                    // 在处理完所有寄存器后更新报警消息
            //                    if (context.AlarmMessages.Length > 0)
            //                    {
            //                        try
            //                        {
            //                            string messageTag = $"meiqi{context.Prefix}";
            //                            hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());
            //                            context.LastAlarmUpdate = DateTime.Now;
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"更新报警消息失败: {ex.Message}");
            //                        }
            //                    }
            //                    else
            //                    {
            //                        // 如果没有报警，清空报警消息
            //                        try
            //                        {
            //                            string messageTag = $"meiqi{context.Prefix}";
            //                            hmi485.Tags[messageTag].Write("无报警");
            //                        }
            //                        catch (Exception ex)
            //                        {
            //                            Debug.WriteLine($"清空报警消息失败: {ex.Message}");
            //                        }
            //                    }

            //                    return true;
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}区响应，报警消息长度: {context.AlarmMessages.Length}");
            //                Debug.WriteLine($"ProcessModbusResponse error: {ex.Message}");
            //            }

            //            return false;
            //        }

            //        private void ProcessRegisterValue(int address, int registerNumber, int value, SerialPortContext context)
            //        {
            //            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //            string displayText = $"{context.Prefix}区（{context.Port.PortName}）第{registerNumber}路数据: 值：{value} 读取时间：{timestamp}，";
            //            string messageTag = $"meiqi{context.Prefix}";
            //            try
            //            {
            //                if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
            //                {
            //                    string statusTag = $"meiqi{context.Prefix}区";
            //                    string valueTag = $"{context.Prefix}{registerNumber}_Value";
            //                    string timeTag = $"{context.Prefix}{registerNumber}_Time";

            //                    // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
            //                    if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
            //                    {
            //                        displayText += $"-> 串口超时，不更新WinCC状态\r\n";
            //                        return;
            //                    }

            //                    hmi485.Tags[statusTag].Write(1);

            //                    // 处理特殊值
            //                    if (value == 0xFAFA) // 探测器没有接
            //                    {
            //                        hmi485.Tags[valueTag].Write(0xFAFA);
            //                        hmi485.Tags[timeTag].Write($"{timestamp} - 没有连接");
            //                        context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器没有连接；");
            //                        displayText += $"-> 探测器没有连接\r\n";
            //                    }
            //                    else if (value == 0xFEFE) // 探测器丢失
            //                    {
            //                        hmi485.Tags[valueTag].Write(0xFEFE);
            //                        hmi485.Tags[timeTag].Write($"{timestamp} - 丢失");
            //                        context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器丢失");
            //                        displayText += $"-> 探测器丢失\r\n";
            //                    }
            //                    else if (value == 0xFCFC) // 探测器故障
            //                    {
            //                        hmi485.Tags[valueTag].Write(0xFCFC);
            //                        hmi485.Tags[timeTag].Write($"{timestamp} - 故障");
            //                        context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber} 探测器故障");
            //                        displayText += $"-> 探测器故障\r\n";
            //                    }
            //                    else // 正常值
            //                    {
            //                        float actualValue = value / 10.0f;
            //                        hmi485.Tags[valueTag].Write(actualValue);
            //                        hmi485.Tags[timeTag].Write(timestamp);
            //                        hmi485.Tags[messageTag].Write($"CO报警地址 {context.Prefix}{registerNumber}  值: {actualValue:F1}ppm");
            //                        // 正常值不添加到报警消息中
            //                        string tagName = $"meiqi{context.Prefix}区";
            //                        hmi485.Tags[tagName].Write(1);
            //                    }
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}{registerNumber}，值: 0x{value:X4}");
            //                displayText += $"! 写入WinCC标签失败: {ex.Message}\r\n";
            //            }

            //            // 更新UI
            //            if (!isPaused && !globalCts.IsCancellationRequested)
            //            {
            //                this.BeginInvoke(new Action(() =>
            //                {
            //                    richTextBox1.AppendText(displayText);
            //                    richTextBox1.ScrollToCaret();
            //                }));
            //            }
            //        }

            //        private string GetModbusErrorDescription(int errorCode)
            //        {
            //            switch (errorCode)
            //            {
            //                case 0x01: return "非法功能";
            //                case 0x02: return "非法数据地址";
            //                case 0x03: return "非法数据值";
            //                case 0x04: return "从站设备故障";
            //                case 0x05: return "确认";
            //                case 0x06: return "从属设备忙";
            //                case 0x08: return "存储奇偶性差错";
            //                case 0x0A: return "不可用网关路径";
            //                case 0x0B: return "网关目标设备响应失败";
            //                default: return "未知错误";
            //            }
            //        }

            //        public static byte[] CRC_16(byte[] data, out byte[] temdata)
            //        {
            //            if (data.Length == 0)
            //                throw new Exception("调用CRC16校验算法,（低字节在前，高字节在后）时发生异常，异常信息：被校验的数组长度为0。");
            //            int xda, xdapoly;
            //            byte i, j, xdabit;
            //            xda = 0xFFFF;
            //            xdapoly = 0xA001;
            //            for (i = 0; i < data.Length; i++)
            //            {
            //                xda ^= data[i];
            //                for (j = 0; j < 8; j++)
            //                {
            //                    xdabit = (byte)(xda & 0x01);
            //                    xda >>= 1;
            //                    if (xdabit == 1)
            //                        xda ^= xdapoly;
            //                }
            //            }
            //            temdata = new byte[2] { (byte)(xda & 0xFF), (byte)(xda >> 8) };
            //            return temdata;
            //        }

            //        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
            //        {
            //            // 停止所有后台操作
            //            globalCts.Cancel();
            //            twoTimer.Stop();

            //            // 停止所有串口轮询并关闭串口
            //            foreach (var context in serialPortContexts)
            //            {
            //                context.IsPolling = false;
            //                context.Cts.Cancel();
            //                context.FrameTimer.Stop();
            //                if (context.Port.IsOpen) context.Port.Close();
            //                context.Dispose();
            //            }

            //            // 更新WinCC状态为关闭 - 确保所有前缀的状态都被更新
            //            try
            //            {
            //                if (!stopWritingToWinCC)
            //                {
            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        try
            //                        {
            //                            //string tagName = $"meiqi{context.Prefix}区";
            //                            //hmi485.Tags[tagName].Write(0);
            //                        }
            //                        catch { /* 忽略单个串口关闭时的异常 */ }
            //                    }
            //                }
            //            }
            //            catch { /* 忽略关闭时的异常 */ }
            //        }

            //        private void richTextBox1_TextChanged(object sender, EventArgs e)
            //        {
            //            if (isPaused || globalCts.IsCancellationRequested) return;

            //            const int maxLine = 1200;
            //            if (richTextBox1.Lines.Length > maxLine)
            //            {
            //                int firstNewLine = richTextBox1.Text.IndexOf('\n');
            //                if (firstNewLine >= 0)
            //                {
            //                    // 保留滚动位置
            //                    int scrollPos = richTextBox1.GetCharIndexFromPosition(new Point(0, 0));
            //                    richTextBox1.Text = richTextBox1.Text.Substring(firstNewLine + 1);
            //                    richTextBox1.ScrollToCaret();
            //                    if (scrollPos > 0)
            //                    {
            //                        richTextBox1.SelectionStart = Math.Min(scrollPos - firstNewLine - 1, richTextBox1.Text.Length);
            //                        richTextBox1.ScrollToCaret();
            //                    }
            //                }
            //            }
            //        }

            //        protected override void Dispose(bool disposing)
            //        {
            //            if (disposing)
            //            {
            //                try
            //                {
            //                    // 取消所有操作
            //                    globalCts?.Cancel();

            //                    // 停止并释放定时器
            //                    twoTimer?.Stop();
            //                    twoTimer?.Dispose();

            //                    // 释放所有串口上下文
            //                    foreach (var context in serialPortContexts)
            //                    {
            //                        context?.Dispose();
            //                    }
            //                    serialPortContexts.Clear();

            //                    // 释放组件资源
            //                    components?.Dispose();
            //                }
            //                catch
            //                {
            //                    // 忽略清理过程中的异常
            //                }
            //                finally
            //                {
            //                    globalCts?.Dispose();
            //                }
            //            }
            //            base.Dispose(disposing);
            //        }

            //        private void button1_Click(object sender, EventArgs e) { isPaused = !isPaused; button1.Text = isPaused ? "继续滚屏显示" : "停止滚屏显示"; }
            //        private void textBox1_TextChanged(object sender, EventArgs e) { }
            //    }
            //}
            #endregion
            #region//0.5
//            using CCHMIRUNTIME;
//            using System;
//            using System.Collections.Concurrent;
//            using System.Collections.Generic;
//            using System.Configuration;
//            using System.Data;
//            using System.Diagnostics;
//            using System.Drawing;
//            using System.IO.Ports;
//            using System.Linq;
//            using System.Text;
//            using System.Threading;
//            using System.Threading.Tasks;
//            using System.Windows.Forms;

//            using System.Runtime.InteropServices;


//namespace meiqibaojing
//    {
//        public partial class Form1 : Form
//        {
//            int RESPONSE_TIMEOUT = int.Parse(ConfigurationManager.AppSettings["RESPONSETIMEOUT"]);
//            private int[] stationAddresses; // 站地址数组
//            private HMIRuntime hmi485 = new CCHMIRUNTIME.HMIRuntime();
//            private bool isPaused = false;
//            private bool stopWritingToWinCC = false;
//            private CancellationTokenSource globalCts = new CancellationTokenSource();
//            private List<SerialPortContext> serialPortContexts = new List<SerialPortContext>();
//            private System.Windows.Forms.Timer twoTimer = new System.Windows.Forms.Timer();
//            // 添加超时状态字典
//            private ConcurrentDictionary<string, bool> portTimeoutStatus = new ConcurrentDictionary<string, bool>();
//            #region//热键
//            [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
//            [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
//            private const int WM_HOTKEY = 0x0312;
//            private const int HOTKEY_ID = 1;  // 热键

//            protected override void OnLoad(EventArgs e)
//            {
//                base.OnLoad(e);

//                // 注册快捷键 (Ctrl + Shift + M)
//                if (!RegisterHotKey(this.Handle, HOTKEY_ID, 0x0002 | 0x0004, (int)Keys.M))
//                {
//                    Debug.WriteLine("热键注册失败");
//                }
//            }

//            protected override void WndProc(ref Message m)
//            {
//                if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
//                {
//                    // 直接调用切换方法
//                    ToggleWindowVisibility();
//                    return;
//                }
//                base.WndProc(ref m);
//            }
//            private void ToggleWindowVisibility()
//            {
//                if (this.InvokeRequired)
//                {
//                    this.Invoke(new Action(ToggleWindowVisibility));
//                    return;
//                }

//                if (this.WindowState == FormWindowState.Minimized || !this.Visible)
//                {
//                    // 显示窗口
//                    this.WindowState = FormWindowState.Normal;
//                    this.Visible = true;
//                    this.ShowInTaskbar = true;
//                    this.Activate();
//                    this.Focus();
//                }
//                else
//                {
//                    // 隐藏窗口
//                    this.WindowState = FormWindowState.Minimized;
//                    this.ShowInTaskbar = false;
//                }
//            }

//            protected override void OnFormClosing(FormClosingEventArgs e)
//            {
//                UnregisterHotKey(this.Handle, HOTKEY_ID);
//                base.OnFormClosing(e);
//            }

//            private void ShowWindow()
//            {
//                if (this.IsDisposed || this.Disposing)
//                    return;

//                if (this.InvokeRequired)
//                {
//                    // 使用 MethodInvoker 确保线程安全
//                    this.BeginInvoke(new MethodInvoker(ShowWindow));
//                    return;
//                }

//                try
//                {
//                    // 确保窗口完全显示
//                    this.Visible = true;
//                    this.ShowInTaskbar = true;
//                    this.WindowState = FormWindowState.Normal;

//                    // 激活窗口
//                    this.Activate();

//                    // 确保窗口获得焦点
//                    this.Focus();

//                    // 将窗口置于前台
//                    this.TopMost = true;
//                    this.TopMost = false;

//                    Debug.WriteLine("窗口显示成功");
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"显示窗口失败: {ex.Message}");
//                }
//            }
//            #endregion
//            private class SerialPortContext : IDisposable
//            {
//                public StringBuilder AlarmMessages { get; } = new StringBuilder();
//                public DateTime LastAlarmUpdate { get; set; } = DateTime.MinValue;
//                public SerialPort Port { get; set; }
//                public string Prefix { get; set; }
//                public int StartRegisterAddress { get; set; }
//                public int RegisterCount { get; set; }
//                public List<byte> ReceiveBuffer { get; } = new List<byte>();
//                public DateTime LastReceivedTime { get; set; } = DateTime.MinValue;
//                public ConcurrentDictionary<int, TaskCompletionSource<bool>> ResponseWaiters { get; } = new ConcurrentDictionary<int, TaskCompletionSource<bool>>();
//                public object SerialLock { get; } = new object();
//                public bool IsPolling { get; set; } = true;
//                public CancellationTokenSource Cts { get; } = new CancellationTokenSource();
//                public int Index { get; set; }  // 串口索引
//                public System.Windows.Forms.Timer FrameTimer { get; set; } // 帧超时定时器
//                public bool IsTimeout { get; set; } = false; // 添加超时状态

//                public void Dispose()
//                {
//                    try
//                    {
//                        Cts?.Cancel();
//                        Cts?.Dispose();
//                        FrameTimer?.Stop();
//                        FrameTimer?.Dispose();

//                        // 确保串口完全释放
//                        if (Port != null)
//                        {
//                            try
//                            {
//                                if (Port.IsOpen)
//                                    Port.Close();
//                                Port.Dispose();
//                                Port = null; // 设置为null防止重复使用
//                            }
//                            catch
//                            {
//                                // 忽略释放异常
//                                Port = null;
//                            }
//                        }
//                    }
//                    catch
//                    {
//                        // 忽略所有异常
//                    }
//                }
//            }// 串口上下文类，管理每个串口的独立状态

//            public Form1()
//            {
//                InitializeComponent();

//                // 初始化系统托盘图标
//                InitializeNotifyIcon();

//                LoadPorts();
//                twoTimer.Interval = 1000;
//                twoTimer.Tick += new EventHandler(WinccConn);
//                twoTimer.Start();
//                this.WindowState = FormWindowState.Minimized;
//                this.ShowInTaskbar = false;
//                SetVisibleCore(false);  // 初始隐藏
//                isPaused = true; // 默认关闭滚屏
//                button1.Text = "打开滚屏"; // 设置初始按钮文本
//            }
//            private void InitializeNotifyIcon()
//            {
//                // 创建NotifyIcon
//                notifyIcon1 = new NotifyIcon();
//                notifyIcon1.Icon = this.Icon; // 使用应用程序图标
//                notifyIcon1.Visible = true;
//                notifyIcon1.Text = "CO报警监控程序";

//                // 创建右键菜单
//                ContextMenuStrip contextMenu = new ContextMenuStrip();
//                ToolStripMenuItem showItem = new ToolStripMenuItem("显示窗口");
//                showItem.Click += (s, e) => ShowWindow();
//                ToolStripMenuItem exitItem = new ToolStripMenuItem("退出程序");
//                exitItem.Click += (s, e) => Application.Exit();

//                contextMenu.Items.Add(showItem);
//                contextMenu.Items.Add(exitItem);

//                notifyIcon1.ContextMenuStrip = contextMenu;

//                // 双击托盘图标显示窗口
//                notifyIcon1.DoubleClick += (s, e) => ShowWindow();
//            }
//            private async void Form1_Load(object sender, EventArgs e)
//            {
//                Task while04 = Task.Run(() => Winccstart());
//                try
//                {
//                    // 初始化所有串口
//                    foreach (var context in serialPortContexts)
//                    {
//                        try
//                        {
//                            context.Port.DataReceived += (s, args) => SerialPort_DataReceived(s, args, context);
//                            context.Port.Open();
//                            context.FrameTimer.Start(); // 启动帧超时检测定时器

//                            // 更新WinCC状态 - 确保每个串口的状态都被更新
//                            string tagName = $"meiqi{context.Prefix}区";
//                            hmi485.Tags[tagName].Write(1);
//                        }
//                        catch (Exception ex)
//                        {
//                            MessageBox.Show($"初始化串口{context.Port.PortName}失败: {ex.Message}");
//                            continue;
//                        }
//                    }

//                    // 启动轮询任务
//                    var pollingTasks = serialPortContexts
//                        .Where(context => context.Port.IsOpen)
//                        .Select(context => StartPollingAsync(context));

//                    await Task.WhenAll(pollingTasks);
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show($"初始化失败: {ex.Message},程序将自动退出");

//                    System.Environment.Exit(System.Environment.ExitCode);
//                    //  Environment.Exit(Environment.ExitCode);
//                }
//            }

//            private void LoadPorts()
//            {
//                try
//                {
//                    int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "200");
//                    string addressesConfig = ConfigurationManager.AppSettings["StationAddresses"];
//                    if (!string.IsNullOrEmpty(addressesConfig))
//                    {
//                        stationAddresses = addressesConfig.Split(',')
//                            .Select(addr => int.Parse(addr.Trim()))
//                            .ToArray();
//                    }
//                    else
//                    {
//                        stationAddresses = new int[] { 1 };
//                    }

//                    string[] portNames = ConfigurationManager.AppSettings["SerialPorts"].Split(';');
//                    int[] baudRates = ConfigurationManager.AppSettings["BaudRates"].Split(';').Select(int.Parse).ToArray();
//                    Parity[] parities = ConfigurationManager.AppSettings["Parities"].Split(';').Select(p => (Parity)Enum.Parse(typeof(Parity), p)).ToArray();
//                    int[] dataBitsArray = ConfigurationManager.AppSettings["DataBitsArray"].Split(';').Select(int.Parse).ToArray();
//                    StopBits[] stopBitsArray = ConfigurationManager.AppSettings["StopBitsArray"].Split(';').Select(s => (StopBits)Enum.Parse(typeof(StopBits), s)).ToArray();
//                    string[] prefixes = ConfigurationManager.AppSettings["AddressPrefixes"].Split(';');
//                    int[] startRegisterAddresses = ConfigurationManager.AppSettings["RegisterAddresses"].Split(';').Select(int.Parse).ToArray();
//                    int[] registerCounts;
//                    string registerCountsConfig = ConfigurationManager.AppSettings["RegisterCounts"];
//                    if (!string.IsNullOrEmpty(registerCountsConfig))
//                    {
//                        registerCounts = registerCountsConfig.Split(';').Select(int.Parse).ToArray();
//                    }
//                    else
//                    {
//                        registerCounts = new int[portNames.Length];
//                        for (int i = 0; i < portNames.Length; i++)
//                        {
//                            registerCounts[i] = 16;
//                        }
//                    }
//                    if (new[] { portNames.Length, baudRates.Length, parities.Length, dataBitsArray.Length, stopBitsArray.Length, prefixes.Length, startRegisterAddresses.Length, registerCounts.Length }
//                       .Distinct().Count() != 1)
//                    {
//                        throw new ConfigurationErrorsException("串口配置数组长度不一致");
//                    }

//                    StringBuilder configText = new StringBuilder("串口配置:\r\n");
//                    for (int i = 0; i < portNames.Length; i++)
//                    {
//                        configText.AppendLine($"串口{i + 1}: {portNames[i]}, {baudRates[i]}, {parities[i]}, {dataBitsArray[i]}, {stopBitsArray[i]}, 区域:{prefixes[i]}, 起始寄存器:{startRegisterAddresses[i]}, 读取数量:{registerCounts[i]}, 轮询间隔:{pollingInterval}ms");// 创建串口上下文
//                        var context = new SerialPortContext
//                        {
//                            Port = new SerialPort(portNames[i], baudRates[i], parities[i], dataBitsArray[i], stopBitsArray[i]),
//                            Prefix = prefixes[i],
//                            StartRegisterAddress = startRegisterAddresses[i],
//                            RegisterCount = registerCounts[i],
//                            Index = i + 1,
//                            FrameTimer = new System.Windows.Forms.Timer { Interval = int.Parse(ConfigurationManager.AppSettings["FrameInterval"] ?? "20") }
//                        };

//                        context.FrameTimer.Tick += (sender, e) => CheckFrameTimeout(context);
//                        serialPortContexts.Add(context);
//                        // 初始化超时状态
//                        portTimeoutStatus[context.Prefix] = false;
//                    }

//                    textBox1.Text = configText.ToString();
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show($"加载配置失败: {ex.Message}");
//                    Environment.Exit(1);
//                }
//            }

//            private void WinccConn(object sender, EventArgs e)
//            {
//                try
//                {
//                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;

//                    foreach (var context in serialPortContexts)
//                    {
//                        try
//                        {
//                            // 添加 null 检查
//                            if (context == null || context.Port == null)
//                            {
//                                continue;
//                            }

//                            // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
//                            if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
//                            {
//                                continue;
//                            }

//                            if (context.Port.IsOpen)
//                            {
//                                string tagName = $"meiqi{context.Prefix}区";
//                                hmi485.Tags[tagName]?.Write(1);
//                            }
//                        }
//                        catch (Exception ex)
//                        {
//                            Debug.WriteLine($"更新串口状态到WinCC失败: {ex.Message}");
//                            // 即使出错也继续处理其他串口
//                        }
//                    }
//                }
//                catch (Exception ex)
//                {
//                    if (stopWritingToWinCC || globalCts.IsCancellationRequested) return;
//                    MessageBox.Show($"与WinCC通讯失败: {ex.Message}，驱动程序将退出。");
//                    Environment.Exit(Environment.ExitCode);
//                }
//            }

//            private async Task Winccstart()
//            {
//                string processName = ConfigurationManager.AppSettings["ProcessName"];

//                while (!globalCts.IsCancellationRequested)
//                {
//                    try
//                    {
//                        bool b = IsProcessStarted(processName);

//                        if (b)
//                        {
//                            stopWritingToWinCC = false;
//                            // 更新所有串口的状态
//                            foreach (var context in serialPortContexts)
//                            {
//                                try
//                                {
//                                    // 添加 null 检查
//                                    if (context == null || context.Port == null)
//                                    {
//                                        continue;
//                                    }

//                                    // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
//                                    if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
//                                    {
//                                        continue;
//                                    }

//                                    if (context.Port.IsOpen)
//                                    {
//                                        string tagName = $"meiqi{context.Prefix}区";
//                                        hmi485.Tags[tagName]?.Write(1);
//                                    }
//                                }
//                                catch (Exception ex)
//                                {
//                                    Debug.WriteLine($"更新串口{context?.Port?.PortName}状态到WinCC失败: {ex.Message}");
//                                }
//                            }
//                        }
//                        else
//                        {
//                            if (InvokeRequired)
//                            {
//                                Invoke(new Action(() =>
//                                {
//                                    DialogResult result = MessageBox.Show(
//                                        $"未找到运行的WinCC实例({processName})，是否退出驱动程序？",
//                                        "WinCC未运行",
//                                        MessageBoxButtons.YesNo,
//                                        MessageBoxIcon.Warning);

//                                    if (result == DialogResult.Yes)
//                                    {
//                                        MessageBox.Show("驱动程序将退出。");
//                                        Environment.Exit(Environment.ExitCode);
//                                    }
//                                    else
//                                    {
//                                        stopWritingToWinCC = true;
//                                        richTextBox1.AppendText($"已停止向WinCC写入标签。WinCC进程({processName})未运行。\r\n");
//                                        richTextBox1.ScrollToCaret();
//                                    }
//                                }));
//                            }

//                            await Task.Delay(10000, globalCts.Token);
//                        }

//                        await Task.Delay(1000, globalCts.Token);
//                    }
//                    catch (OperationCanceledException)
//                    {
//                        // 任务被取消，正常退出
//                        return;
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine($"WinCC进程检测错误: {ex.Message}");
//                        await Task.Delay(5000, globalCts.Token);
//                    }
//                }
//            }

//            bool IsProcessStarted(string processName1)
//            {
//                try
//                {
//                    Process[] temp = Process.GetProcessesByName(processName1);
//                    return temp.Length > 0;
//                }
//                catch
//                {
//                    return false;
//                }
//            }

//            private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e, SerialPortContext context)
//            {
//                try
//                {
//                    lock (context.SerialLock)
//                    {
//                        int bytesToRead = context.Port.BytesToRead;
//                        if (bytesToRead <= 0) return;

//                        byte[] buffer = new byte[bytesToRead];
//                        context.Port.Read(buffer, 0, bytesToRead);

//                        // 更新最后接收时间
//                        context.LastReceivedTime = DateTime.Now;

//                        // 将数据添加到缓冲区
//                        context.ReceiveBuffer.AddRange(buffer);

//                        // 重置帧超时定时器
//                        context.FrameTimer.Stop();
//                        context.FrameTimer.Start();

//                        // 尝试处理缓冲区中的完整帧
//                        ProcessReceivedData(context);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"SerialPort_DataReceived error: {ex.Message}");
//                    this.BeginInvoke(new Action(() =>
//                    {
//                        richTextBox1.AppendText($"串口{context.Port.PortName}接收数据错误: {ex.Message}\r\n");
//                        richTextBox1.ScrollToCaret();
//                    }));
//                }
//            }

//            private void ProcessReceivedData(SerialPortContext context)
//            {
//                try
//                {
//                    if (context.ReceiveBuffer.Count == 0) return;

//                    // 查找可能的完整帧
//                    while (context.ReceiveBuffer.Count >= 5) // 最小帧长度
//                    {
//                        // 检查是否为错误响应帧（5字节）
//                        if (context.ReceiveBuffer.Count >= 5 &&
//                            (context.ReceiveBuffer[1] & 0x80) != 0)
//                        {
//                            // 错误响应帧长度为5字节
//                            if (context.ReceiveBuffer.Count >= 5)
//                            {
//                                byte[] errorFrame = context.ReceiveBuffer.Take(5).ToArray();
//                                ProcessCompleteFrame(errorFrame, context);
//                                context.ReceiveBuffer.RemoveRange(0, 5);
//                                continue;
//                            }
//                        }

//                        // 检查正常响应帧
//                        if (context.ReceiveBuffer[1] == 0x03) // 读取保持寄存器
//                        {
//                            // 正常响应帧长度 = 地址(1) + 功能码(1) + 字节数(1) + 数据(n*2) + CRC(2)
//                            int expectedLength = 3 + context.ReceiveBuffer[2] + 2;

//                            if (context.ReceiveBuffer.Count >= expectedLength)
//                            {
//                                byte[] completeFrame = context.ReceiveBuffer.Take(expectedLength).ToArray();
//                                ProcessCompleteFrame(completeFrame, context);
//                                context.ReceiveBuffer.RemoveRange(0, expectedLength);
//                                continue;
//                            }
//                            else
//                            {
//                                // 帧不完整，等待更多数据
//                                break;
//                            }
//                        }
//                        else
//                        {
//                            // 未知帧类型，尝试查找下一个可能的帧头
//                            int nextStart = FindNextFrameStart(context.ReceiveBuffer);
//                            if (nextStart > 0)
//                            {
//                                // 移除无效数据
//                                context.ReceiveBuffer.RemoveRange(0, nextStart);
//                                continue;
//                            }
//                            else
//                            {
//                                // 没有找到有效帧头，清空缓冲区
//                                context.ReceiveBuffer.Clear();
//                                break;
//                            }
//                        }
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"ProcessReceivedData error: {ex.Message}");
//                    context.ReceiveBuffer.Clear();
//                }
//            }

//            private int FindNextFrameStart(List<byte> buffer)
//            {
//                for (int i = 1; i < buffer.Count; i++)
//                {
//                    // 有效的Modbus地址范围是1-247
//                    if (buffer[i] >= 1 && buffer[i] <= 247)
//                    {
//                        // 检查是否是有效的功能码
//                        if (i + 1 < buffer.Count)
//                        {
//                            byte functionCode = buffer[i + 1];
//                            if (functionCode == 0x03 || (functionCode & 0x80) != 0)
//                            {
//                                return i;
//                            }
//                        }
//                    }
//                }
//                return -1;
//            }

//            // 处理完整的帧
//            private void ProcessCompleteFrame(byte[] frame, SerialPortContext context)
//            {
//                try
//                {
//                    // 显示接收到的原始数据
//                    string rawDataHex = BitConverter.ToString(frame).Replace("-", " ");
//                    string timestamp = DateTime.Now.ToString("mm:ss");
//                    string displayText = $"{context.Prefix}区（{context.Port.PortName}）收到{frame.Length}字节: {rawDataHex} -> {timestamp} \r\n";

//                    // 写入WinCC接收报文变量
//                    try
//                    {
//                        if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
//                        {
//                            string messageTag = $"meiqi{context.Prefix}1";
//                            string receiveMessage = $"{timestamp}=>{rawDataHex}";
//                            hmi485.Tags[messageTag].Write(receiveMessage);
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine($"写入WinCC接收报文失败: {ex.Message}");
//                    }

//                    // 处理Modbus响应
//                    bool isModbusResponse = ProcessModbusResponse(frame, context);

//                    if (!isModbusResponse)
//                    {
//                        displayText += $"-> 非标准Modbus响应或CRC校验失败\r\n";
//                    }

//                    // 更新UI显示
//                    if (!isPaused && !globalCts.IsCancellationRequested)
//                    {
//                        this.BeginInvoke(new Action(() =>
//                        {
//                            richTextBox1.AppendText(displayText);
//                            richTextBox1.ScrollToCaret();
//                        }));
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"ProcessCompleteFrame error: {ex.Message}");
//                }
//            }

//            // 帧超时检测
//            private void CheckFrameTimeout(SerialPortContext context)
//            {
//                lock (context.SerialLock)
//                {
//                    context.FrameTimer.Stop();

//                    if (context.ReceiveBuffer.Count > 0)
//                    {
//                        // 帧超时，处理缓冲区中可能的不完整数据
//                        string timeoutText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 帧超时，丢弃{context.ReceiveBuffer.Count}字节不完整数据\r\n";

//                        if (!isPaused && !globalCts.IsCancellationRequested)
//                        {
//                            this.BeginInvoke(new Action(() =>
//                            {
//                                richTextBox1.AppendText(timeoutText);
//                                richTextBox1.ScrollToCaret();
//                            }));
//                        }

//                        context.ReceiveBuffer.Clear();
//                    }

//                    context.FrameTimer.Start();
//                }
//            }

//            // 在 StartPollingAsync 方法中修改串口状态检查逻辑
//            private async Task StartPollingAsync(SerialPortContext context)
//            {
//                int currentIndex = 0;
//                int pollingInterval = int.Parse(ConfigurationManager.AppSettings["PollingInterval"] ?? "100");

//                while (context.IsPolling && !context.Cts.Token.IsCancellationRequested && !globalCts.IsCancellationRequested)
//                {
//                    if (stationAddresses == null || stationAddresses.Length == 0)
//                    {
//                        await Task.Delay(100, context.Cts.Token);
//                        continue;
//                    }

//                    int address = stationAddresses[currentIndex];
//                    currentIndex = (currentIndex + 1) % stationAddresses.Length;

//                    try
//                    {
//                        // 发送请求
//                        byte[] request = BuildModbusRequest(address, context.StartRegisterAddress, context.RegisterCount, context);

//                        // 获取该地址的等待器
//                        var tcs = new TaskCompletionSource<bool>();
//                        context.ResponseWaiters[address] = tcs;

//                        using (var ctsTimeout = new CancellationTokenSource(RESPONSE_TIMEOUT))
//                        {
//                            var responseTask = tcs.Task;
//                            var timeoutTask = Task.Delay(Timeout.Infinite, ctsTimeout.Token);

//                            var completedTask = await Task.WhenAny(responseTask, timeoutTask);

//                            if (completedTask == responseTask)
//                            {
//                                // 收到响应，取消超时
//                                ctsTimeout.Cancel();
//                                // 清除超时状态
//                                portTimeoutStatus[context.Prefix] = false;
//                                context.IsTimeout = false;
//                            }
//                            else
//                            {
//                                // 超时处理
//                                context.ResponseWaiters.TryRemove(address, out _);
//                                HandleResponseTimeout(context, address);
//                            }
//                        }
//                    }
//                    catch (OperationCanceledException)
//                    {
//                        // 取消操作，继续下一个
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine($"轮询错误: {ex.Message}");
//                        await Task.Delay(1000, context.Cts.Token);
//                    }

//                    // 添加轮询间隔，避免过于频繁
//                    await Task.Delay(pollingInterval, context.Cts.Token);
//                }
//            }

//            // 处理响应超时

//            private void HandleResponseTimeout(SerialPortContext context, int address)
//            {
//                if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
//                {
//                    try
//                    {
//                        // 清空并设置报警消息
//                        context.AlarmMessages.Clear();
//                        string timeoutMessage = $"{context.Prefix}区({context.Port.PortName}) 报警器分线箱离线";
//                        context.AlarmMessages.AppendLine(timeoutMessage);

//                        // 写入超时信息到WinCC报文变量
//                        string timestamp = DateTime.Now.ToString("mm:ss");
//                        string timeoutInfo = $"{timestamp}=>{context.Prefix}区 串口{context.Port.PortName}无响应";
//                        string messageTag = $"meiqi{context.Prefix}1";
//                        hmi485.Tags[messageTag].Write(timeoutInfo);

//                        // 更新报警消息标签
//                        hmi485.Tags[$"meiqi{context.Prefix}"].Write(context.AlarmMessages.ToString());

//                        // 更新所有寄存器的超时状态
//                        for (int i = 1; i <= context.RegisterCount; i++)
//                        {
//                            hmi485.Tags[$"{context.Prefix}{i}_Value"].Write(999999); // 使用特殊值表示超时
//                            hmi485.Tags[$"{context.Prefix}{i}_Time"].Write(DateTime.Now.ToString("离线时间：yyyy-MM-dd HH:mm:ss.fff"));
//                        }

//                        // 记录超时日志
//                        Debug.WriteLine($"地址 {address} 响应超时，已更新WinCC标签");

//                        if (!isPaused && !globalCts.IsCancellationRequested)
//                        {
//                            this.BeginInvoke(new Action(() =>
//                            {
//                                richTextBox1.AppendText($"{timestamp} {context.Prefix}区{context.Port.PortName}: 报警器分线箱响应超时，已更新WinCC标签\r\n");
//                                richTextBox1.ScrollToCaret();
//                            }));
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine($"写入WinCC超时状态失败: {ex.Message}");

//                        this.BeginInvoke(new Action(() =>
//                        {
//                            richTextBox1.AppendText($"写入WinCC超时状态失败: {ex.Message}\r\n");
//                            richTextBox1.ScrollToCaret();
//                        }));
//                    }
//                }
//            }

//            private byte[] BuildModbusRequest(int address, int startRegister, int registerCount, SerialPortContext context)
//            {
//                byte[] request = new byte[8];
//                request[0] = (byte)address;          // 设备地址
//                request[1] = 0x03;                   // 功能码
//                request[2] = (byte)(startRegister >> 8);  // 寄存器地址高字节
//                request[3] = (byte)(startRegister & 0xFF);// 寄存器地址低字节
//                request[4] = (byte)(registerCount >> 8);  // 读取数量高字节
//                request[5] = (byte)(registerCount & 0xFF);// 读取数量低字节

//                // 计算CRC校验
//                byte[] crc;
//                CRC_16(request.Take(6).ToArray(), out crc);
//                request[6] = crc[0];
//                request[7] = crc[1];

//                // 将字节数组转换为十六进制字符串
//                string hexString = BitConverter.ToString(request).Replace("-", " ");
//                string timestamp = DateTime.Now.ToString("mm:ss");
//                string requestMessage = $"{timestamp}=>{hexString}";

//                // 写入WinCC变量
//                try
//                {
//                    if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
//                    {
//                        string messageTag = $"meiqi{context.Prefix}1";
//                        hmi485.Tags[messageTag].Write(requestMessage);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"写入WinCC发送报文失败: {ex.Message}");
//                }

//                // 显示完整的请求帧
//                if (!isPaused && !globalCts.IsCancellationRequested)
//                {
//                    this.BeginInvoke(new Action(() =>
//                    {
//                        richTextBox1.AppendText($"{context.Prefix}区{context.Port.PortName}请求: {hexString}\r\n");
//                        richTextBox1.ScrollToCaret();
//                    }));
//                }

//                // 实际发送数据
//                lock (context.SerialLock)
//                {
//                    if (context.Port.IsOpen && !globalCts.IsCancellationRequested)
//                    {
//                        try
//                        {
//                            context.Port.Write(request, 0, request.Length);
//                        }
//                        catch (Exception ex)
//                        {
//                            Debug.WriteLine($"发送数据失败: {ex.Message}");
//                        }
//                    }
//                }

//                return request;
//            }

//            private bool ProcessModbusResponse(byte[] response, SerialPortContext context)
//            {
//                if (response.Length < 5) return false;

//                try
//                {
//                    // 在处理响应前清空报警消息
//                    context.AlarmMessages.Clear();

//                    // 验证CRC
//                    byte[] crc;
//                    CRC_16(response.Take(response.Length - 2).ToArray(), out crc);
//                    if (crc[0] != response[response.Length - 2] || crc[1] != response[response.Length - 1])
//                    {
//                        Debug.WriteLine("CRC校验失败");
//                        return false;
//                    }

//                    int address = response[0]; // 设备地址
//                    int functionCode = response[1];

//                    // 检查错误响应
//                    if ((functionCode & 0x80) != 0)
//                    {
//                        int errorCode = response[2];
//                        string errorMessage = GetModbusErrorDescription(errorCode);
//                        string errorText = $"{DateTime.Now:HH:mm:ss.fff} [{context.Prefix}-{context.Port.PortName}] 地址{address}错误响应: {errorMessage}\r\n";

//                        // 写入错误信息到WinCC
//                        try
//                        {
//                            if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
//                            {
//                                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
//                                string errorInfo = $"[错误] {timestamp} 地址{address}: {errorMessage}";
//                                string messageTag = $"meiqi{context.Prefix}1";
//                                hmi485.Tags[messageTag].Write(errorInfo);
//                            }
//                        }
//                        catch (Exception ex)
//                        {
//                            Debug.WriteLine($"写入WinCC错误信息失败: {ex.Message}");
//                        }

//                        this.BeginInvoke(new Action(() =>
//                        {
//                            richTextBox1.AppendText(errorText);
//                            richTextBox1.ScrollToCaret();
//                        }));

//                        return true;
//                    }

//                    if (functionCode == 0x03) // 读取保持寄存器响应
//                    {
//                        int byteCount = response[2];
//                        if (byteCount != context.RegisterCount * 2) return false;

//                        // 通知等待器
//                        if (context.ResponseWaiters.TryRemove(address, out var tcs))
//                        {
//                            tcs.TrySetResult(true);
//                        }

//                        // 清除超时状态
//                        portTimeoutStatus[context.Prefix] = false;
//                        context.IsTimeout = false;

//                        // 解析所有寄存器值并更新到WinCC
//                        for (int i = 0; i < context.RegisterCount; i++)
//                        {
//                            int registerIndex = 3 + i * 2; // 数据起始位置
//                            int value = (response[registerIndex] << 8) | response[registerIndex + 1];
//                            int registerNumber = i + 1; // 寄存器编号从1开始

//                            // 处理特殊值和更新WinCC
//                            ProcessRegisterValue(address, registerNumber, value, context);
//                        }

//                        // 在处理完所有寄存器后更新报警消息
//                        if (context.AlarmMessages.Length > 0)
//                        {
//                            try
//                            {
//                                string messageTag = $"meiqi{context.Prefix}";
//                                hmi485.Tags[messageTag].Write(context.AlarmMessages.ToString());
//                                context.LastAlarmUpdate = DateTime.Now;
//                            }
//                            catch (Exception ex)
//                            {
//                                Debug.WriteLine($"更新报警消息失败: {ex.Message}");
//                            }
//                        }
//                        else
//                        {
//                            // 如果没有报警，清空报警消息
//                            try
//                            {
//                                string messageTag = $"meiqi{context.Prefix}";
//                                hmi485.Tags[messageTag].Write("无报警");
//                            }
//                            catch (Exception ex)
//                            {
//                                Debug.WriteLine($"清空报警消息失败: {ex.Message}");
//                            }
//                        }

//                        return true;
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}区响应，报警消息长度: {context.AlarmMessages.Length}");
//                    Debug.WriteLine($"ProcessModbusResponse error: {ex.Message}");
//                }

//                return false;
//            }

//            private void ProcessRegisterValue(int address, int registerNumber, int value, SerialPortContext context)
//            {
//                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
//                string displayText = $"{context.Prefix}区（{context.Port.PortName}）第{registerNumber}路数据: 值：{value} 读取时间：{timestamp}，";
//                string messageTag = $"meiqi{context.Prefix}";
//                try
//                {
//                    if (!stopWritingToWinCC && !globalCts.IsCancellationRequested)
//                    {
//                        string statusTag = $"meiqi{context.Prefix}区";
//                        string valueTag = $"{context.Prefix}{registerNumber}_Value";
//                        string timeTag = $"{context.Prefix}{registerNumber}_Time";

//                        // 检查该串口是否处于超时状态，如果是则不更新WinCC状态
//                        if (portTimeoutStatus.ContainsKey(context.Prefix) && portTimeoutStatus[context.Prefix])
//                        {
//                            displayText += $"-> 串口超时，不更新WinCC状态\r\n";
//                            return;
//                        }

//                        hmi485.Tags[statusTag].Write(1);

//                        // 处理特殊值
//                        if (value == 0xFAFA) // 探测器没有接
//                        {
//                            hmi485.Tags[valueTag].Write(0xFAFA);
//                            hmi485.Tags[timeTag].Write($"{timestamp} - 没有连接");
//                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber}没有连接；");
//                            displayText += $"->没有连接\r\n";
//                        }
//                        else if (value == 0xFEFE) // 探测器丢失
//                        {
//                            hmi485.Tags[valueTag].Write(0xFEFE);
//                            hmi485.Tags[timeTag].Write($"{timestamp} - 丢失");
//                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber}丢失");
//                            displayText += $"->丢失\r\n";
//                        }
//                        else if (value == 0xFCFC) // 探测器故障
//                        {
//                            hmi485.Tags[valueTag].Write(0xFCFC);
//                            hmi485.Tags[timeTag].Write($"{timestamp} - 故障");
//                            context.AlarmMessages.AppendLine($"{context.Prefix}{registerNumber}故障");
//                            displayText += $"->故障\r\n";
//                        }
//                        else // 正常值
//                        {
//                            float actualValue = value / 10.0f;
//                            hmi485.Tags[valueTag].Write(actualValue);
//                            hmi485.Tags[timeTag].Write(timestamp);
//                            hmi485.Tags[messageTag].Write($"CO报警地址 {context.Prefix}{registerNumber}  值: {actualValue:F1}ppm");
//                            // 正常值不添加到报警消息中
//                            string tagName = $"meiqi{context.Prefix}区";
//                            hmi485.Tags[tagName].Write(1);
//                        }
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} 处理 {context.Prefix}{registerNumber}，值: 0x{value:X4}");
//                    displayText += $"! 写入WinCC标签失败: {ex.Message}\r\n";
//                }

//                // 更新UI
//                if (!isPaused && !globalCts.IsCancellationRequested)
//                {
//                    this.BeginInvoke(new Action(() =>
//                    {
//                        richTextBox1.AppendText(displayText);
//                        richTextBox1.ScrollToCaret();
//                    }));
//                }
//            }

//            private string GetModbusErrorDescription(int errorCode)
//            {
//                switch (errorCode)
//                {
//                    case 0x01: return "非法功能";
//                    case 0x02: return "非法数据地址";
//                    case 0x03: return "非法数据值";
//                    case 0x04: return "从站设备故障";
//                    case 0x05: return "确认";
//                    case 0x06: return "从属设备忙";
//                    case 0x08: return "存储奇偶性差错";
//                    case 0x0A: return "不可用网关路径";
//                    case 0x0B: return "网关目标设备响应失败";
//                    default: return "未知错误";
//                }
//            }

//            public static byte[] CRC_16(byte[] data, out byte[] temdata)
//            {
//                if (data.Length == 0)
//                    throw new Exception("调用CRC16校验算法,（低字节在前，高字节在后）时发生异常，异常信息：被校验的数组长度为0。");
//                int xda, xdapoly;
//                byte i, j, xdabit;
//                xda = 0xFFFF;
//                xdapoly = 0xA001;
//                for (i = 0; i < data.Length; i++)
//                {
//                    xda ^= data[i];
//                    for (j = 0; j < 8; j++)
//                    {
//                        xdabit = (byte)(xda & 0x01);
//                        xda >>= 1;
//                        if (xdabit == 1)
//                            xda ^= xdapoly;
//                    }
//                }
//                temdata = new byte[2] { (byte)(xda & 0xFF), (byte)(xda >> 8) };
//                return temdata;
//            }

//            private void Form1_FormClosing(object sender, FormClosingEventArgs e)
//            {
//                // 停止所有后台操作
//                globalCts.Cancel();
//                twoTimer.Stop();

//                // 停止所有串口轮询并关闭串口
//                foreach (var context in serialPortContexts)
//                {
//                    context.IsPolling = false;
//                    context.Cts.Cancel();
//                    context.FrameTimer.Stop();
//                    if (context.Port.IsOpen) context.Port.Close();
//                    context.Dispose();
//                }

//                // 更新WinCC状态为关闭 - 确保所有前缀的状态都被更新
//                try
//                {
//                    if (!stopWritingToWinCC)
//                    {
//                        foreach (var context in serialPortContexts)
//                        {
//                            try
//                            {
//                                //string tagName = $"meiqi{context.Prefix}区";
//                                //hmi485.Tags[tagName].Write(0);
//                            }
//                            catch { /* 忽略单个串口关闭时的异常 */ }
//                        }
//                    }
//                }
//                catch
//                {

//                }

//            }
//            private async void CleanupApplication()
//            {
//                try
//                {
//                    // 1. 设置全局取消标记
//                    if (!globalCts.IsCancellationRequested)
//                    {
//                        globalCts.Cancel();
//                    }

//                    // 2. 停止WinCC连接定时器
//                    twoTimer?.Stop();

//                    // 3. 停止所有串口轮询任务并等待完成
//                    var cleanupTasks = new List<Task>();

//                    foreach (var context in serialPortContexts)
//                    {
//                        if (context != null)
//                        {
//                            // 停止轮询
//                            context.IsPolling = false;

//                            // 取消该上下文的令牌
//                            if (context.Cts != null && !context.Cts.IsCancellationRequested)
//                            {
//                                context.Cts.Cancel();
//                            }

//                            // 停止帧超时定时器
//                            context.FrameTimer?.Stop();

//                            // 关闭串口
//                            if (context.Port != null && context.Port.IsOpen)
//                            {
//                                try
//                                {
//                                    context.Port.Close();
//                                }
//                                catch (Exception ex)
//                                {
//                                    Debug.WriteLine($"关闭串口{context.Port.PortName}时出错: {ex.Message}");
//                                }
//                            }
//                        }
//                    }

//                    // 4. 等待一小段时间让所有任务完成清理
//                    try
//                    {
//                        await Task.Delay(500, CancellationToken.None); // 给予500ms的清理时间
//                    }
//                    catch { }

//                    // 5. 强制释放所有资源
//                    foreach (var context in serialPortContexts)
//                    {
//                        if (context != null)
//                        {
//                            try
//                            {
//                                context.Dispose();
//                            }
//                            catch (Exception ex)
//                            {
//                                Debug.WriteLine($"释放串口上下文资源时出错: {ex.Message}");
//                            }
//                        }
//                    }

//                    // 6. 清理WinCC连接
//                    try
//                    {
//                        // 更新WinCC状态为关闭
//                        if (!stopWritingToWinCC)
//                        {
//                            foreach (var context in serialPortContexts)
//                            {
//                                if (context != null)
//                                {
//                                    try
//                                    {
//                                        string tagName = $"meiqi{context.Prefix}区";
//                                        hmi485.Tags[tagName]?.Write(0);
//                                    }
//                                    catch
//                                    {
//                                        // 忽略单个串口关闭时的异常 
//                                    }
//                                }
//                            }
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        Debug.WriteLine($"清理WinCC连接时出错: {ex.Message}");
//                    }

//                    // 7. 释放定时器资源
//                    try
//                    {
//                        twoTimer?.Stop();
//                        twoTimer?.Dispose();
//                    }
//                    catch { }

//                    // 8. 清理集合
//                    serialPortContexts?.Clear();

//                    // 9. 释放全局CancellationTokenSource
//                    try
//                    {
//                        globalCts?.Dispose();
//                    }
//                    catch { }

//                    Debug.WriteLine("应用程序清理完成");
//                }
//                catch (Exception ex)
//                {
//                    Debug.WriteLine($"清理应用程序时发生错误: {ex.Message}");
//                }
//            }

//            private void richTextBox1_TextChanged(object sender, EventArgs e)
//            {
//                if (isPaused || globalCts.IsCancellationRequested) return;

//                const int maxLine = 1200;
//                if (richTextBox1.Lines.Length > maxLine)
//                {
//                    int firstNewLine = richTextBox1.Text.IndexOf('\n');
//                    if (firstNewLine >= 0)
//                    {
//                        // 保留滚动位置
//                        int scrollPos = richTextBox1.GetCharIndexFromPosition(new Point(0, 0));
//                        richTextBox1.Text = richTextBox1.Text.Substring(firstNewLine + 1);
//                        richTextBox1.ScrollToCaret();
//                        if (scrollPos > 0)
//                        {
//                            richTextBox1.SelectionStart = Math.Min(scrollPos - firstNewLine - 1, richTextBox1.Text.Length);
//                            richTextBox1.ScrollToCaret();
//                        }
//                    }
//                }
//            }

//            protected override void Dispose(bool disposing)
//            {
//                if (disposing)
//                {
//                    // 清理托盘图标
//                    if (notifyIcon1 != null)
//                    {
//                        notifyIcon1.Visible = false;
//                        notifyIcon1.Dispose();
//                    }

//                    try
//                    {
//                        // 取消所有操作
//                        globalCts?.Cancel();

//                        // 停止并释放定时器
//                        twoTimer?.Stop();
//                        twoTimer?.Dispose();

//                        // 释放所有串口上下文
//                        foreach (var context in serialPortContexts)
//                        {
//                            context?.Dispose();
//                        }
//                        serialPortContexts.Clear();

//                        // 释放组件资源
//                        components?.Dispose();
//                    }
//                    catch
//                    {
//                        // 忽略清理过程中的异常
//                    }
//                    finally
//                    {
//                        globalCts?.Dispose();
//                    }
//                }
//                base.Dispose(disposing);
//            }

//            private void button1_Click(object sender, EventArgs e)
//            {
//                isPaused = !isPaused;
//                button1.Text = isPaused ? "打开滚屏" : "停止滚屏";
//            }
//            private void textBox1_TextChanged(object sender, EventArgs e) { }

//            private void button2_Click(object sender, EventArgs e)
//            {
//                //// 只最小化，不完全隐藏
//                //this.WindowState = FormWindowState.Minimized;
//                //this.ShowInTaskbar = false;
//                //// 隐藏到系统托盘
//                this.WindowState = FormWindowState.Minimized;
//                this.ShowInTaskbar = false;
//                this.Visible = false;
//                // 移除 SetVisibleCore(false) 调用


//            }

//            private void button3_Click(object sender, EventArgs e)
//            {
//                string helpText = @"软件使用说明：
//    wincc需要建立的内部变量：
//    1.A1_Time（文本变量 16 位字符集）
//    2.A1_Value（32-位浮点数 IEEE 754）
//    3.meiqiA（文本16位字符集）
//    4.meiqiA1（文本16位字符集）
//    5.meiqiA区（int32，程序通讯指示器）
//    6.meiqiA区断线（bool，变量5每秒+1放入此变量用于判断状态）
//    其他使用说明： 
//    1.功能：配置文件可以设置串口数量及通讯参数
//    2.注意事项：配置文件中  PdlRt 为wincc的后台进程名称
//    3.调试快捷健：Ctrl+Shift+M
//    4.wincc脚本建立范本(位于bin\Debug文件夹)
    
//                        版本：2.0
//                        更新时间：2025.8.23 RNI-ELT2";

//                MessageBox.Show(helpText, "Wincc与串口通讯程序使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
//            }

//            private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
//            {

//            }
//        }
//    }
    #endregion



}
    }
}
