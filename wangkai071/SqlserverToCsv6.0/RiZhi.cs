using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace DatabasePRG
{
    public class RiZhi : ILogger
    {
        /// 日志服务实现类 - 处理所有日志记录
        /// 特点：线程安全、自动滚动、异常防护,根据不同方法自动加前缀

        private static RiZhi _instance;
        public static RiZhi Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new RiZhi();
                }
                return _instance;
            }
        }

        // 实例方法
        private readonly Dictionary<string, string> _levelPrefixes = new Dictionary<string, string>
        {
            { "ERROR", "❌ " }, { "SUCCESS", "✅ " }, { "INFO", "ℹ️  " },  { "WARNING", "⚠️  " }
        }; //添加日志前缀
        public void LogMessage(string message, params TextBox[] logTargets)
        {
            AppendLog(message, "", logTargets); //3个参数：日志输出内容，前缀符号，日志数组
        }
        public void LogError(string message, params TextBox[] logTargets)
        {
            AppendLog(message, "ERROR", logTargets);
        }
        public void LogSuccess(string message, params TextBox[] logTargets)
        {
            AppendLog(message, "SUCCESS", logTargets);
        }
        public void LogInfo(string message, params TextBox[] logTargets)
        {
            AppendLog(message, "INFO", logTargets);  // 修复：使用 "INFO"
        }
        public void LogWarning(string message, params TextBox[] logTargets)
        {
            AppendLog(message, "WARNING", logTargets);  // 修复：使用 "WARNING"
        }

        // 静态方法 - 可以直接使用 RiZhi.LogSuccess(...) 调用
        public static void Log(string message, params TextBox[] logTargets)
        {
            Instance.LogMessage(message, logTargets);
        }
        public static void Error(string message, params TextBox[] logTargets)
        {
            Instance.LogError(message, logTargets);
        }
        public static void Success(string message, params TextBox[] logTargets)
        {
            Instance.LogSuccess(message, logTargets);
        }
        public static void Info(string message, params TextBox[] logTargets)
        {
            Instance.LogInfo(message, logTargets);
        }
        public static void Warning(string message, params TextBox[] logTargets)
        {
            Instance.LogWarning(message, logTargets);
        }

        private void AppendLog(string message, string level, params TextBox[] logTargets)
        {
            if (logTargets == null || logTargets.Length == 0) return;

            string prefix = _levelPrefixes.ContainsKey(level) ? _levelPrefixes[level] : ""; //前缀符号
            string timestamp = DateTime.Now.ToString("HH:mm:ss"); //日期
            string logEntry = $"[{timestamp}] {prefix}{message}\r\n"; //日期+前缀+内容

            // 遍历所有目标文本框
            foreach (TextBox tb in logTargets) //遍历所有文本框数组
            {
                if (tb == null || tb.IsDisposed) continue;

                // 线程安全调用
                try
                {
                    if (tb.InvokeRequired)
                    {
                        tb.Invoke(new Action(() => SafeAppendText(tb, logEntry)));
                    }
                    else
                    {
                        SafeAppendText(tb, logEntry);
                    }
                }
                catch (ObjectDisposedException) { /* 控件已销毁，静默忽略 */ }
                catch (InvalidOperationException) { /* 控件状态异常，静默忽略 */ }
            }
        }/// 核心日志追加方法
        private void SafeAppendText(TextBox textBox, string text)
        {
            try
            {
                // 限制日志长度防止内存溢出（可选）
                if (textBox.TextLength > 100000)
                {
                    textBox.Clear();
                    textBox.AppendText("[日志已清空 - 长度限制]\r\n");
                }

                textBox.AppendText(text);
                textBox.SelectionStart = textBox.Text.Length;
                textBox.ScrollToCaret();
            }
            catch { /* 最终兜底，确保不中断主流程 */ }
        }/// （处理滚动和异常）安全追加文本到TextBox
    }
}
//_log.LogSuccess("应用程序已启动", _logList.ToArray());
//RiZhi.Warning("应用程序已启动...", _logList.ToArray()); 以上两种写法为什么感觉2好一些