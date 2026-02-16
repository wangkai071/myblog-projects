using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseCreator01
{
    public class ApplicationLogger
    {
        // 定义日志事件，允许UI订阅
        public event Action<string> OnLogMessage;

        // 线程安全的日志记录方法
        public void Log(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logEntry = $"[{timestamp}] {message}\r\n";

            // 通过事件通知所有订阅者
            OnLogMessage?.Invoke(logEntry);
        }

        // 可选：添加不同级别的日志方法
        public void LogError(string message) => Log($"ERROR: {message}");
        public void LogWarning(string message) => Log($"WARNING: {message}");
    }
}
