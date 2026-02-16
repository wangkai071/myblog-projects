using System;
using System.Collections.Generic;

namespace DatabasePRG.Export
{
    /// <summary>
    /// CSV 导出参数
    /// 导出参数类 (ExcelParams)，用于导出为 CSV 文件
    /// 设计原因：将导出所需的所有参数封装在一个对象中，便于传递和管理。避免方法签名过长，并且可以方便地扩展新的参数。
    /// </summary>
    public class ExcelParams
    {
        public string ConnectionString { get; set; }
        public List<string> SelectedTables { get; set; } = new List<string>();
        public string FilePath { get; set; }
        public bool EnableTimeFilter { get; set; }
        public string TimeColumnDisplay { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public IProgressReporter ProgressReporter { get; set; }
        public Dictionary<string, List<string>> TableColumns { get; set; } = new Dictionary<string, List<string>>();
        public Dictionary<string, Dictionary<string, string>> ColumnDataTypes { get; set; } = new Dictionary<string, Dictionary<string, string>>();
    }

    /// <summary>
    /// CSV 导出结果类 (ExcelExportResult)
    /// 设计原因：封装导出操作的结果，包括成功状态、导出表数、行数、文件路径和消息。便于统一返回结果，并且可以方便地扩展更多返回信息
    /// </summary>
    public class ExcelExportResult
    {
        public bool Success { get; set; }
        public int ExportedTables { get; set; }
        public long ExportedRows { get; set; }
        public string FilePath { get; set; }
        public string Message { get; set; }
        public TimeSpan ElapsedTime { get; set; }
        public Exception Error { get; set; }
    }

    /// <summary>
    /// 进度报告接口
    /// 设计意义：
    //参数封装：将散乱的参数（原来通过多个控件获取）封装成单一对象

    //降低耦合：导出器不需要知道具体的UI控件，只需要参数对象

    //易于扩展：新增参数只需修改这个类，不影响调用方

    //类型安全：避免字符串拼接和类型转换错误
    // 原代码：业务逻辑与UI紧密耦合
//    progressBar1.Value = percentage;
//lblProgress.Text = status;
//Application.DoEvents(); // 强制刷新（危险！）

//// 新设计：通过接口调用，不知道具体UI
//_progressReporter.ReportProgress(percentage, status);
    /// </summary>
    public interface IProgressReporter
    {
        void ReportProgress(int percentage, string status);
        void LogMessage(string message, LogLevel level = LogLevel.Info);
    }

    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Info,
        Success,
        Warning,
        Error
    }
}