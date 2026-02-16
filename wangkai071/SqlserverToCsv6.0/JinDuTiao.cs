using System;
using System.Windows.Forms;

namespace DatabasePRG.Export
{
    /// <summary>
    /// 窗体进度报告器（适配 WinForms UI）
    /// </summary>
    public class JinDuTiao : IProgressReporter
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _progressLabel;
        private readonly Label _percentageLabel;
        private readonly ILogger _logger;
        private readonly TextBox[] _logTargets;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="progressBar">进度条控件</param>
        /// <param name="progressLabel">进度文本标签</param>
        /// <param name="percentageLabel">百分比标签</param>
        /// <param name="logger">日志记录器</param>
        /// <param name="logTargets">日志输出目标文本框</param>
        public JinDuTiao(ProgressBar progressBar, Label progressLabel,
                                   Label percentageLabel, ILogger logger, params TextBox[] logTargets)
        {
            _progressBar = progressBar ?? throw new ArgumentNullException(nameof(progressBar));
            _progressLabel = progressLabel ?? throw new ArgumentNullException(nameof(progressLabel));
            _percentageLabel = percentageLabel ?? throw new ArgumentNullException(nameof(percentageLabel));
            _logger = logger;
            _logTargets = logTargets;
        }

        /// <summary>
        /// 报告进度
        /// </summary>
        /// <param name="percentage">进度百分比（0-100）</param>
        /// <param name="status">状态信息</param>
        public void ReportProgress(int percentage, string status)
        {
            if (_progressBar.InvokeRequired)
            {
                _progressBar.Invoke(new Action(() => ReportProgress(percentage, status)));
                return;
            }

            // 确保百分比在有效范围内
            percentage = Math.Max(0, Math.Min(100, percentage));

            // 更新UI控件
            _progressBar.Value = percentage;
            _progressLabel.Text = status ?? string.Empty;
            _percentageLabel.Text = $"{percentage}%";

            // 可选：强制UI刷新
            Application.DoEvents();
        }

        /// <summary>
        /// 记录日志消息
        /// </summary>
        /// <param name="message">消息内容</param>
        /// <param name="level">日志级别</param>
        public void LogMessage(string message, LogLevel level = LogLevel.Info)
        {
            // 如果提供了日志记录器，使用它记录
            if (_logger != null && _logTargets != null)
            {
                switch (level)
                {
                    case LogLevel.Success:
                        _logger.LogSuccess(message, _logTargets);
                        break;
                    case LogLevel.Warning:
                        _logger.LogWarning(message, _logTargets);
                        break;
                    case LogLevel.Error:
                        _logger.LogError(message, _logTargets);
                        break;
                    default:
                        _logger.LogInfo(message, _logTargets);
                        break;
                }
            }

            // 也可以直接在进度标签中显示重要消息
            if (level == LogLevel.Error || level == LogLevel.Success)
            {
                if (_progressLabel.InvokeRequired)
                {
                    _progressLabel.Invoke(new Action(() => _progressLabel.Text = message));
                }
                else
                {
                    _progressLabel.Text = message;
                }
            }
        }

        /// <summary>
        /// 重置进度显示
        /// </summary>
        public void ResetProgress()
        {
            if (_progressBar.InvokeRequired)
            {
                _progressBar.Invoke(new Action(ResetProgress));
                return;
            }

            _progressBar.Value = 0;
            _progressLabel.Text = "就绪";
            _percentageLabel.Text = "0%";
        }

        /// <summary>
        /// 显示完成状态
        /// </summary>
        public void ShowComplete(string message = "导出完成")
        {
            ReportProgress(100, message);
        }

        /// <summary>
        /// 显示错误状态
        /// </summary>
        public void ShowError(string errorMessage)
        {
            ReportProgress(0, $"错误: {errorMessage}");
            LogMessage(errorMessage, LogLevel.Error);
        }
    }
}