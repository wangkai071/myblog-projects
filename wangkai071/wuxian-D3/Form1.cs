using CCHMIRUNTIME;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace wuxian
{
    public partial class Form1 : Form
    {
        private const byte PACKET_START1   = 0x09;   private const int PACKET_MIN_SIZE = 14;
        private const byte PACKET_START2   = 0xAF;   private const int PACKET_MAX_SIZE = 28;
        private const byte PACKET_END      = 0x16;
        
        HMIRuntime hmi485 = new CCHMIRUNTIME.HMIRuntime();

        private static string   portName = ConfigurationManager.AppSettings["SerialPort"];
        private static int      baudRate = int.Parse(ConfigurationManager.AppSettings["BaudRate"]);
        private static Parity   parity   = (Parity)Enum.Parse(typeof(Parity), ConfigurationManager.AppSettings["Parity"]);
        private static int      dataBits = int.Parse(ConfigurationManager.AppSettings["DataBits"]);
        private static StopBits stopBits = (StopBits)Enum.Parse(typeof(StopBits), ConfigurationManager.AppSettings["StopBits"]);
        SerialPort port1    = new SerialPort  (portName,   baudRate,   parity,   dataBits,   stopBits);
        string displayText  = $"默认串口配置: {portName}, {baudRate}, {parity}, {dataBits}, {stopBits}";

        private object     serialLock    = new object();
        private List<byte> receiveBuffer = new List<byte>();
        private readonly ConcurrentDictionary<int, DateTime> nodeLastUpdateTimes = new ConcurrentDictionary<int, DateTime>();
        private System.Threading.Timer nodeMonitorTimer;// 使用ConcurrentDictionary存储节点最后更新时间

        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            SetVisibleCore(false);  // 初始隐藏
            textBox1.Text = displayText;
        }
        [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        private const int WM_HOTKEY = 0x0312; private const int HOTKEY_ID = 1;  // 热键
        protected override void OnLoad(EventArgs e)
        {
            // 注册快捷键 (Ctrl + Shift + H)
            RegisterHotKey(this.Handle, HOTKEY_ID, 0x0002 | 0x0004, (int)Keys.W);
            base.OnLoad(e);
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                ShowWindow();  // 快捷键触发时显示窗口
            }
            base.WndProc(ref m);
        }
        private void ShowWindow()
        {
            SetVisibleCore(true);   // 设为可见
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;
            this.Activate();  // 激活窗口
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosing(e);
        }// 清理热键
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                port1.Close();
                port1.DataReceived += SerialPort_white;
                port1.Open();
                nodeMonitorTimer = new System.Threading.Timer(CheckNodeStatus, null, TimeSpan.Zero, TimeSpan.FromMinutes(1));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败: {ex.Message}");
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            nodeMonitorTimer?.Dispose();
            if (port1.IsOpen)
            {
                port1.Close();
            }
        }
        public static byte[] CRC_16(byte[] data, out byte[] temdata)
        {
            if (data.Length == 0)
                throw new Exception("调用CRC16校验算法,（低字节在前，高字节在后）时发生异常，异常信息：被校验的数组长度为0。");
            int xda, xdapoly;
            byte i, j, xdabit;
            xda = 0xFFFF;
            xdapoly = 0xA001;
            for (i = 0; i < data.Length; i++)
            {
                xda ^= data[i];
                for (j = 0; j < 8; j++)
                {
                    xdabit = (byte)(xda & 0x01);
                    xda >>= 1;
                    if (xdabit == 1)
                        xda ^= xdapoly;
                }
            }
            temdata = new byte[2] { (byte)(xda & 0xFF), (byte)(xda >> 8) };
            return temdata;
        }
        public static string byteArrayToHexString(byte[] data)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                builder.Append(string.Format("{0:X2} ", data[i]));
            }
            return builder.ToString().Trim();
        }
        public static byte[] hexStringToByteArray(string data)
        {
            string[] chars = data.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            byte[] returnBytes = new byte[chars.Length];
            for (int i = 0; i < chars.Length; i++)
            {
                returnBytes[i] = Convert.ToByte(chars[i], 16);
            }
            return returnBytes;
        }
        private void SerialPort_white(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                hmi485.Tags[("wuxian通讯程序开启状态")].Write(1);

                lock (serialLock)
                {
                    int bytesToRead = port1.BytesToRead;
                    if (bytesToRead <= 0) return;
                    byte[] buffer = new byte[bytesToRead];
                    port1.Read(buffer, 0, bytesToRead);
                    receiveBuffer.AddRange(buffer);
                    byte[] byteArray = receiveBuffer.ToArray();
                    this.BeginInvoke(new Action(() =>
                    {
                        richTextBox1.AppendText("\r\n" + byteArrayToHexString(buffer));
                        richTextBox1.ScrollToCaret();
                    }));
                    ProcessBuffer();
                }
            }
            catch (Exception ex)
            {
                port1.DiscardInBuffer();
                Debug.WriteLine($"SerialPort_white error: {ex.Message}");
            }
        }
        private void ProcessBuffer()
        {
               while (receiveBuffer.Count >= PACKET_MIN_SIZE)
            {
                int startIndex = -1;
                for (int i = 0; i <= receiveBuffer.Count - 2; i++)
                {
                    if (receiveBuffer[i] == PACKET_START1 && receiveBuffer[i + 1] == PACKET_START2)
                    {
                        startIndex = i;
                        break;
                    }
                }

                if (startIndex == -1)
                {
                    if (receiveBuffer.Count > 1)
                    {
                        receiveBuffer.RemoveRange(0, receiveBuffer.Count - 1);
                    }
                    return;
                }

                if (startIndex > 0)
                {
                    receiveBuffer.RemoveRange(0, startIndex);
                    startIndex = 0;
                }

                int endIndex = -1;
                for (int i = startIndex + 2; i < receiveBuffer.Count; i++)
                {
                    if (receiveBuffer[i] == PACKET_END)
                    {
                        endIndex = i;
                        break;
                    }
                }

                if (endIndex == -1)
                {
                    return;
                }

                int packetLength = endIndex - startIndex + 1;

                if (packetLength != PACKET_MIN_SIZE && packetLength != PACKET_MAX_SIZE)
                {
                    // 提取即将被丢弃的数据包内容
                    byte[] invalidPacket = receiveBuffer.GetRange(0, packetLength).ToArray();
                    // 将字节数组转换为十六进制字符串（更易读的格式）
                    string hexString = BitConverter.ToString(invalidPacket).Replace("-", " ");
                    hmi485.Tags["xiupu3"].Write($"非法字符串 {hexString} ");
                    receiveBuffer.RemoveRange(0, packetLength);
                    continue;
                }

                byte[] packet = receiveBuffer.GetRange(0, packetLength).ToArray();
                receiveBuffer.RemoveRange(0, packetLength);
                ProcessPacket(packet);
            }
        }
        private void ProcessPacket(byte[] respBytes)
        {
            try
            {
                if (respBytes.Length == 14)
                {
                    ProcessSinglePacket(respBytes);
                }
                else if (respBytes.Length == 28)
                {
                    byte[] firstPacket = new byte[14];
                    byte[] secondPacket = new byte[14];
                    Array.Copy(respBytes, 0, firstPacket, 0, 14);
                    Array.Copy(respBytes, 14, secondPacket, 0, 14);
                    ProcessSinglePacket(firstPacket);
                    ProcessSinglePacket(secondPacket);
                }
               
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ProcessPacket error: {ex.Message}");
            }
        }
        private void ProcessSinglePacket(byte[] respBytes)
        {
            if (respBytes.Length != 14) return;

            byte[] respBytes11 = new byte[9];
            Array.Copy(respBytes, 2, respBytes11, 0, 9);

            byte[] respBytesCRC = new byte[2];
            CRC_16(respBytes11, out respBytesCRC);

            if (respBytes[0] == PACKET_START1 && respBytes[1] == PACKET_START2 &&  respBytes[13] == PACKET_END && respBytes[11] == respBytesCRC[0] && respBytes[12] == respBytesCRC[1])
            {
                int bianhao = Convert.ToInt32(string.Concat(new[]{4,3,2,1}.Select(i => Convert.ToString(respBytes11[i], 2).PadLeft(8, '0'))),2);
                int wendu   = Convert.ToInt32(string.Concat(new[]{ 6, 5  }.Select(i => Convert.ToString(respBytes11[i], 2).PadLeft(8, '0'))),2);
                
                string elapsedTime118 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                // 更新节点最后接收时间
                nodeLastUpdateTimes[bianhao] = DateTime.Now;
                hmi485.Tags[("" + bianhao)].Write(wendu);
                string bianhaotime = ("" + bianhao) + "time";
                hmi485.Tags[bianhaotime].Write(elapsedTime118);
                hmi485.Tags["xiupu1"].Write($"{portName}发来数据：");
                hmi485.Tags["xiupu2"].Write($"无线节点编号 {bianhao} / 值 {wendu}");
            }
            else
            {
                Debug.WriteLine("Invalid packet received");
            }
        }
        private void CheckNodeStatus(object state) // 使用定时器代替线程睡眠进行节点监控
        {
            try
            {
                DateTime now = DateTime.Now;
                foreach (var nodeId in nodeLastUpdateTimes.Keys.ToList())
                {
                    if (nodeLastUpdateTimes.TryGetValue(nodeId, out DateTime lastUpdate))
                    {
                        if ((now - lastUpdate).TotalMinutes > 15)
                        {
                            hmi485.Tags[$"{nodeId}"].Write(99999);
                            string timeduqu = lastUpdate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                            hmi485.Tags[$"{nodeId}time"].Write($"{timeduqu} - 最后通讯时间");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CheckNodeStatus error: {ex.Message}");
            }
        }
        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {
            int maxLine = 45;
            if (richTextBox1.Lines.Length > maxLine)
            {
                richTextBox1.Text = richTextBox1.Text.Substring(richTextBox1.Lines[0].Length + 1);
            }
        }
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }
    }
}