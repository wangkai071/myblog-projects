using System.Windows.Forms;

namespace DatabasePRG
{
    
    public interface ILogger/// 日志服务接口 - 定义统一日志记录写法
    {
        void LogMessage(string message, params TextBox[] logTargets);
        void LogError(string message, params TextBox[] logTargets);
        void LogWarning(string message, params TextBox[] logTargets);
        void LogInfo(string message, params TextBox[] logTargets);
        void LogSuccess(string message, params TextBox[] logTargets);
    }
}