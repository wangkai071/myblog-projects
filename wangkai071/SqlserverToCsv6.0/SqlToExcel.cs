using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Globalization;
using DatabasePRG;
//设计意义（分层次架构）：
//第一层：接口层（与外界交互）
//csharp
//public ExcelExportResult Export(ExcelParams parameters)
//职责：接收参数，返回结果

//好处：调用简单，输入输出明确

//第二层：验证层（保障健壮性）
//csharp
//private void ValidateParameters(ExcelParams parameters)
//{
//    if (string.IsNullOrEmpty(parameters.ConnectionString))
//        throw new ArgumentException("连接字符串不能为空");
//    // ... 其他验证
//}
//职责：确保输入有效

//好处：提前发现问题，避免执行到一半出错

//第三层：核心逻辑层（真正的业务）
//csharp
//private ExcelExportResult ExportToExcelInternal(ExcelParams parameters)
//{
//    // 统计总行数
//    long totalRows = CalculateTotalRows(parameters);

//    // 导出数据
//    var exportStats = ExportToExcelFile(parameters, totalRows);

//    // 构建结果
//    return new ExcelExportResult { /* ... */ };
//}
//职责：实现具体业务逻辑

//好处：逻辑集中，便于维护和测试

//第四层：数据访问层（与数据库交互）
//csharp
//private long GetTableRowCount(string connectionString, string tableName)
//{
//    using (SqlConnection conn = new SqlConnection(connectionString))
//    {
//        conn.Open();
//        // 执行SQL查询
//    }
//}
//职责：处理数据库操作

//好处：数据库相关代码集中，易于修改和优化

namespace DatabasePRG.Export
{
    /// <summary>
    /// SQL Server 数据导出到 CSV 的导出器
    /// </summary>
    public class SqlToExcel : IDisposable
    {
        /// <summary>单文件最大行数，超过则自动拆分为多个 CSV 并并行导出</summary>
        private const int CsvChunkRowLimit = 5000;

        #region 私有字段
        private readonly ILogger _logger;
        private BackgroundWorker _backgroundWorker;
        private bool _disposed = false;
        private bool _isExporting = false;
        private ExcelExportResult _exportResult;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly object _lockObject = new object();
        #endregion

        #region 属性
        /// <summary>
        /// 是否正在导出
        /// </summary>
        public bool IsExporting => _isExporting;

        /// <summary>
        /// 是否支持取消
        /// </summary>
        public bool SupportsCancellation => true;

        /// <summary>
        /// 导出进度（0-100）
        /// </summary>
        public int ProgressPercentage { get; private set; }
        #endregion

        #region 事件
        /// <summary>
        /// 导出进度变化事件
        /// </summary>
        public event EventHandler<ProgressChangedEventArgs> ProgressChanged;

        /// <summary>
        /// 导出完成事件
        /// </summary>
        public event EventHandler<ExportCompletedEventArgs> ExportCompleted;
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public SqlToExcel(ILogger logger = null)
        {
            _logger = logger;
            InitializeBackgroundWorker();
        }

        private void InitializeBackgroundWorker()
        {
            _backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            _backgroundWorker.DoWork += BackgroundWorker_DoWork;
            _backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }
        #endregion

        #region 公共方法
        /// <summary>
        /// 异步导出数据到 CSV
        /// </summary>
        /// <param name="parameters">导出参数</param>
        /// <returns>导出结果</returns>
        public async Task<ExcelExportResult> ExportAsync(ExcelParams parameters)
        {
            return await Task.Run(() => Export(parameters));
        }

        /// <summary>
        /// 导出数据到 CSV
        /// </summary>
        /// <param name="parameters">导出参数</param>
        /// <returns>导出结果</returns>
        public ExcelExportResult Export(ExcelParams parameters)
        {
            if (parameters == null)
                throw new ArgumentNullException(nameof(parameters));

            if (_isExporting)
                throw new InvalidOperationException("导出器正忙，请等待当前操作完成");

            lock (_lockObject)
            {
                _isExporting = true;
                _exportResult = null;
                _cancellationTokenSource = new CancellationTokenSource();
            }

            var stopwatch = Stopwatch.StartNew();

            try
            {
                // 验证参数
                ValidateParameters(parameters);

                // 运行后台工作线程
                if (_backgroundWorker.IsBusy)
                {
                    throw new InvalidOperationException("后台工作线程正忙");
                }

                _backgroundWorker.RunWorkerAsync(parameters);

                // 等待导出完成或取消
                while (_isExporting && !_backgroundWorker.CancellationPending)
                {
                    Thread.Sleep(100);
                }

                // 返回结果
                if (_exportResult != null)
                {
                    _exportResult.ElapsedTime = stopwatch.Elapsed;
                    return _exportResult;
                }
                else
                {
                    return new ExcelExportResult
                    {
                        Success = false,
                        Message = "导出操作未完成",
                        ElapsedTime = stopwatch.Elapsed
                    };
                }
            }
            catch (Exception ex)
            {
                var result = new ExcelExportResult
                {
                    Success = false,
                    Message = $"导出失败：{ex.Message}",
                    Error = ex,
                    ElapsedTime = stopwatch.Elapsed
                };

                OnExportCompleted(result);
                return result;
            }
            finally
            {
                stopwatch.Stop();
            }
        }

        /// <summary>
        /// 取消导出操作
        /// </summary>
        public void CancelExport()
        {
            if (_backgroundWorker != null && _backgroundWorker.IsBusy)
            {
                _backgroundWorker.CancelAsync();
                _cancellationTokenSource?.Cancel();
                _isExporting = false;
            }
        }
        #endregion

        #region 后台工作线程事件处理
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var parameters = e.Argument as ExcelParams;

            try
            {
                // 执行导出
                var result = ExecuteExport(parameters, worker, _cancellationTokenSource.Token);
                e.Result = result;
            }
            catch (OperationCanceledException)
            {
                e.Cancel = true;
            }
            catch (Exception ex)
            {
                e.Result = new ExcelExportResult
                {
                    Success = false,
                    Message = $"导出过程中出错：{ex.Message}",
                    Error = ex
                };
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressPercentage = e.ProgressPercentage;
            OnProgressChanged(e);
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lock (_lockObject)
            {
                _isExporting = false;

                if (e.Error != null)
                {
                    _exportResult = new ExcelExportResult
                    {
                        Success = false,
                        Message = $"导出失败：{e.Error.Message}",
                        Error = e.Error
                    };
                }
                else if (e.Cancelled)
                {
                    _exportResult = new ExcelExportResult
                    {
                        Success = false,
                        Message = "导出已取消"
                    };
                }
                else
                {
                    _exportResult = e.Result as ExcelExportResult;
                }

                // 触发完成事件
                OnExportCompleted(_exportResult);
            }
        }
        #endregion

        #region 核心导出逻辑
        private ExcelExportResult ExecuteExport(ExcelParams parameters,
                                              BackgroundWorker worker,
                                              CancellationToken cancellationToken)
        {
            var stopwatch = Stopwatch.StartNew();
            var result = new ExcelExportResult();

            try
            {
                // 1. 解析时间列
                var timeColumnInfo = ParseTimeColumn(parameters.TimeColumnDisplay);

                // 2. 统计总行数
                worker.ReportProgress(0, new ProgressData
                {
                    Status = "正在统计总条目数...",
                    CurrentTable = null,
                    TableProgress = 0,
                    OverallProgress = 0
                });

                var rowCountInfo = CalculateTotalRows(parameters, timeColumnInfo, worker, cancellationToken);
                if (cancellationToken.IsCancellationRequested)
                {
                    return new ExcelExportResult { Success = false, Message = "导出已取消" };
                }

                // 3. 执行导出
                worker.ReportProgress(5, new ProgressData
                {
                    Status = $"统计完成，总条目数: {rowCountInfo.TotalRows}，开始导出...",
                    CurrentTable = null,
                    TableProgress = 0,
                    OverallProgress = 5
                });

                var exportStats = ExportToCsvFiles(parameters, timeColumnInfo, rowCountInfo, worker, cancellationToken);
                if (cancellationToken.IsCancellationRequested)
                {
                    return new ExcelExportResult { Success = false, Message = "导出已取消" };
                }

                // 4. 构建结果
                stopwatch.Stop();

                result.Success = true;
                result.ExportedTables = exportStats.TablesExported;
                result.ExportedRows = exportStats.RowsExported;
                // 单表导出时返回实际 CSV 文件路径，多表时返回所在目录
                if (exportStats.TablesExported == 1)
                {
                    result.FilePath = parameters.FilePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                        ? Path.GetFullPath(parameters.FilePath)
                        : Path.GetFullPath(Path.ChangeExtension(parameters.FilePath, ".csv"));
                }
                else
                {
                    result.FilePath = string.IsNullOrEmpty(Path.GetDirectoryName(parameters.FilePath))
                        ? Path.GetFullPath(".")
                        : Path.GetFullPath(Path.GetDirectoryName(parameters.FilePath));
                }
                result.ElapsedTime = stopwatch.Elapsed;
                result.Message = $"成功导出 {exportStats.TablesExported} 个表为 CSV 文件，共 {exportStats.RowsExported} 条数据，耗时：{stopwatch.Elapsed:hh\\:mm\\:ss}";
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.Message = $"导出失败：{ex.Message}";
                result.Error = ex;
                result.ElapsedTime = stopwatch.Elapsed;
            }

            return result;
        }

        private TimeColumnInfo ParseTimeColumn(string timeColumnDisplay)
        {
            if (string.IsNullOrEmpty(timeColumnDisplay))
                return null;

            var info = new TimeColumnInfo();

            int parenIndex = timeColumnDisplay.LastIndexOf('(');
            if (parenIndex > 0)
            {
                info.ColumnName = timeColumnDisplay.Substring(0, parenIndex).Trim();
                info.DataType = timeColumnDisplay.Substring(parenIndex + 1).TrimEnd(')').Trim();
            }
            else
            {
                info.ColumnName = timeColumnDisplay;
            }

            return info;
        }

        private RowCountInfo CalculateTotalRows(ExcelParams parameters,
                                              TimeColumnInfo timeColumnInfo,
                                              BackgroundWorker worker,
                                              CancellationToken cancellationToken)
        {
            long totalRows = 0;
            var tableRowCounts = new Dictionary<string, long>();
            int totalTables = parameters.SelectedTables.Count;

            for (int i = 0; i < totalTables; i++)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                string tableName = parameters.SelectedTables[i];

                int progress = (int)((i * 100.0) / Math.Max(1, totalTables));
                worker.ReportProgress(progress, new ProgressData
                {
                    Status = $"正在统计表 '{tableName}' 的条目数... ({i + 1}/{totalTables})",
                    CurrentTable = tableName,
                    TableProgress = 0,
                    OverallProgress = progress
                });

                long rowCount = GetTableRowCount(
                    parameters.ConnectionString,
                    tableName,
                    parameters.EnableTimeFilter ? timeColumnInfo?.ColumnName : null,
                    parameters.EnableTimeFilter ? timeColumnInfo?.DataType : null,
                    parameters.EnableTimeFilter ? parameters.StartTime : (DateTime?)null,
                    parameters.EnableTimeFilter ? parameters.EndTime : (DateTime?)null);

                tableRowCounts[tableName] = rowCount;
                totalRows += rowCount;
            }

            return new RowCountInfo
            {
                TotalRows = totalRows,
                TableRowCounts = tableRowCounts
            };
        }

        private ExportStats ExportToCsvFiles(ExcelParams parameters,
                                            TimeColumnInfo timeColumnInfo,
                                            RowCountInfo rowCountInfo,
                                            BackgroundWorker worker,
                                            CancellationToken cancellationToken)
        {
            var stats = new ExportStats();
            string directory = Path.GetDirectoryName(parameters.FilePath);
            string baseName = Path.GetFileNameWithoutExtension(parameters.FilePath);
            if (string.IsNullOrEmpty(directory)) directory = ".";

            int totalTables = parameters.SelectedTables.Count;
            long exportedRows = 0;
            int exportedTables = 0;

            for (int i = 0; i < totalTables; i++)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                string tableName = parameters.SelectedTables[i];
                long currentTableRows = rowCountInfo.TableRowCounts.ContainsKey(tableName)
                    ? rowCountInfo.TableRowCounts[tableName]
                    : 0;

                int tableStartProgress = CalculateOverallProgress(exportedRows, rowCountInfo.TotalRows, exportedTables, totalTables);
                worker.ReportProgress(tableStartProgress, new ProgressData
                {
                    Status = $"正在导出表: {tableName} (共 {currentTableRows} 条)",
                    CurrentTable = tableName,
                    TableProgress = 0,
                    OverallProgress = tableStartProgress
                });

                try
                {
                    List<string> columns = null;
                    if (parameters.TableColumns != null && parameters.TableColumns.ContainsKey(tableName))
                    {
                        columns = parameters.TableColumns[tableName];
                    }

                    string safeTableFileName = GetSafeFileName(tableName);
                    string tableBaseName = totalTables == 1 ? baseName : $"{baseName}_{safeTableFileName}";
                    long rowCount = currentTableRows;
                    bool useChunkedExport = rowCount > CsvChunkRowLimit;
                    long rowsWrittenThisTable = rowCount;

                    if (!useChunkedExport)
                    {
                        // 不超过 10000 条：单次查询，单文件导出
                        var dataTable = GetDataFromSqlServer(
                            parameters.ConnectionString,
                            tableName,
                            columns,
                            parameters.EnableTimeFilter ? timeColumnInfo?.ColumnName : null,
                            parameters.EnableTimeFilter ? timeColumnInfo?.DataType : null,
                            parameters.EnableTimeFilter ? parameters.StartTime : (DateTime?)null,
                            parameters.EnableTimeFilter ? parameters.EndTime : (DateTime?)null);
                        rowCount = dataTable.Rows.Count;
                        rowsWrittenThisTable = rowCount;
                        string csvFilePath = totalTables == 1
                            ? parameters.FilePath
                            : Path.Combine(directory, $"{tableBaseName}.csv");
                        if (!csvFilePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                            csvFilePath = Path.ChangeExtension(csvFilePath, ".csv");

                        using (var writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
                        {
                            writer.Write('\uFEFF');
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                if (j > 0) writer.Write(',');
                                writer.Write(EscapeCsvField(dataTable.Columns[j].ColumnName));
                            }
                            writer.WriteLine();
                            for (int j = 0; j < dataTable.Rows.Count; j++)
                            {
                                if (cancellationToken.IsCancellationRequested) break;
                                for (int k = 0; k < dataTable.Columns.Count; k++)
                                {
                                    if (k > 0) writer.Write(',');
                                    writer.Write(EscapeCsvField(FormatCellForCsv(dataTable.Rows[j][k])));
                                }
                                writer.WriteLine();
                                if (j % 100 == 0 || j == dataTable.Rows.Count - 1)
                                {
                                    rowsExportedSoFar = exportedRows + j + 1;
                                    worker.ReportProgress(
                                        CalculateOverallProgress(rowsExportedSoFar, rowCountInfo.TotalRows, exportedTables, totalTables),
                                        new ProgressData
                                        {
                                            Status = $"正在导出表 '{tableName}' ({j + 1}/{dataTable.Rows.Count})",
                                            CurrentTable = tableName,
                                            TableProgress = (int)((j + 1) * 100.0 / dataTable.Rows.Count),
                                            OverallProgress = CalculateOverallProgress(rowsExportedSoFar, rowCountInfo.TotalRows, exportedTables, totalTables)
                                        });
                                }
                            }
                        }
                    }
                    else
                    {
                        // 超过 10000 条：一次全表查询 + 并行写多个 CSV（OFFSET 分片查询在 SQL Server 上更慢，故不采用）
                        var dataTable = GetDataFromSqlServer(
                            parameters.ConnectionString,
                            tableName,
                            columns,
                            parameters.EnableTimeFilter ? timeColumnInfo?.ColumnName : null,
                            parameters.EnableTimeFilter ? timeColumnInfo?.DataType : null,
                            parameters.EnableTimeFilter ? parameters.StartTime : (DateTime?)null,
                            parameters.EnableTimeFilter ? parameters.EndTime : (DateTime?)null);
                        rowCount = dataTable.Rows.Count;
                        var chunkList = new List<CsvChunkInfo>();
                        for (int start = 0; start < rowCount && !cancellationToken.IsCancellationRequested; start += CsvChunkRowLimit)
                        {
                            int count = Math.Min(CsvChunkRowLimit, (int)rowCount - start);
                            int partIndex = chunkList.Count + 1;
                            string chunkFilePath = Path.Combine(directory, $"{tableBaseName}_{partIndex}.csv");
                            chunkList.Add(new CsvChunkInfo { StartIndex = start, RowCount = count, FilePath = chunkFilePath });
                        }
                        rowsWrittenThisTable = 0;
                        var options = new ParallelOptions { CancellationToken = cancellationToken };
                        try
                        {
                            Parallel.ForEach(chunkList, options, (chunk, state) =>
                            {
                                if (cancellationToken.IsCancellationRequested) { state.Stop(); return; }
                                WriteChunkToCsv(dataTable, chunk.StartIndex, chunk.RowCount, chunk.FilePath);
                                long written = Interlocked.Add(ref rowsWrittenThisTable, chunk.RowCount);
                                worker.ReportProgress(CalculateOverallProgress(exportedRows + written, rowCountInfo.TotalRows, exportedTables, totalTables), new ProgressData
                                {
                                    Status = $"正在并行导出表 '{tableName}' (已写 {written}/{rowCount} 条)",
                                    CurrentTable = tableName,
                                    TableProgress = (int)(written * 100.0 / rowCount),
                                    OverallProgress = CalculateOverallProgress(exportedRows + written, rowCountInfo.TotalRows, exportedTables, totalTables)
                                });
                            });
                        }
                        catch (OperationCanceledException) { }
                    }

                    exportedRows += (int)rowsWrittenThisTable;
                    exportedTables++;
                    stats.RowsExported = exportedRows;
                    stats.TablesExported = exportedTables;

                    int finalProgress = CalculateOverallProgress(exportedRows, rowCountInfo.TotalRows, exportedTables, totalTables);
                    worker.ReportProgress(finalProgress, new ProgressData
                    {
                        Status = useChunkedExport
                            ? $"表 '{tableName}' 导出完成 ({rowsWrittenThisTable} 条，已拆分为多个 CSV 并行导出)"
                            : $"表 '{tableName}' 导出完成 ({rowsWrittenThisTable} 条)",
                        CurrentTable = null,
                        TableProgress = 100,
                        OverallProgress = finalProgress
                    });
                }
                catch (Exception ex)
                {
                    worker.ReportProgress(ProgressPercentage, new ProgressData
                    {
                        Status = $"表 '{tableName}' 导出失败: {ex.Message}",
                        CurrentTable = tableName,
                        TableProgress = 0,
                        OverallProgress = ProgressPercentage
                    });
                }
            }

            return stats;
        }

        /// <summary>将 DataTable 的指定行范围写入单个 CSV 文件（含 BOM 与表头）</summary>
        private static void WriteChunkToCsv(DataTable dataTable, int startIndex, int rowCount, string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                writer.Write('\uFEFF');
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    if (c > 0) writer.Write(',');
                    writer.Write(EscapeCsvField(dataTable.Columns[c].ColumnName));
                }
                writer.WriteLine();
                int end = startIndex + rowCount;
                for (int r = startIndex; r < end; r++)
                {
                    for (int c = 0; c < dataTable.Columns.Count; c++)
                    {
                        if (c > 0) writer.Write(',');
                        writer.Write(EscapeCsvField(FormatCellForCsv(dataTable.Rows[r][c])));
                    }
                    writer.WriteLine();
                }
            }
        }

        /// <summary>CSV 字段转义（逗号、引号、换行用双引号包裹并转义内部引号）</summary>
        private static string EscapeCsvField(string value)
        {
            if (value == null) return "";
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        private static string FormatCellForCsv(object value)
        {
            if (value == null || value == DBNull.Value) return "";
            if (value is DateTime dt)
                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            return value.ToString();
        }

        /// <summary>生成可用于文件名的表名（去掉非法字符）</summary>
        private static string GetSafeFileName(string tableName)
        {
            if (string.IsNullOrEmpty(tableName)) return "Table";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(tableName.Length);
            foreach (char c in tableName)
            {
                if (Array.IndexOf(invalid, c) >= 0)
                    sb.Append('_');
                else
                    sb.Append(c);
            }
            string name = sb.ToString();
            if (name.Length > 50) name = name.Substring(0, 50);
            return string.IsNullOrEmpty(name) ? "Table" : name;
        }
        #endregion

        #region 数据库操作方法（从原窗体迁移）
        private long GetTableRowCount(string connectionString, string tableName,
                                     string timeColumnName = null, string timeColumnDataType = null,
                                     DateTime? startTime = null, DateTime? endTime = null)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // 安全处理表名
                string safeTableName = GetSafeTableName(tableName);

                // 构建查询
                string query = $"SELECT COUNT(*) FROM {safeTableName}";

                // 添加时间过滤
                if (!string.IsNullOrEmpty(timeColumnName) && startTime.HasValue && endTime.HasValue)
                {
                    string safeTimeColumnName = GetSafeColumnName(timeColumnName);

                    if (!string.IsNullOrEmpty(timeColumnDataType))
                    {
                        if (timeColumnDataType.StartsWith("BIGINT") || timeColumnDataType.StartsWith("INT"))
                        {
                            long startTimeValue = long.Parse(startTime.Value.ToString("yyyyMMddHHmm"));
                            long endTimeValue = long.Parse(endTime.Value.ToString("yyyyMMddHHmm"));

                            query += $" WHERE [{safeTimeColumnName}] >= @StartTime AND [{safeTimeColumnName}] <= @EndTime";

                            using (var command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@StartTime", startTimeValue);
                                command.Parameters.AddWithValue("@EndTime", endTimeValue);
                                return Convert.ToInt64(command.ExecuteScalar());
                            }
                        }
                    }

                    query += $" WHERE [{safeTimeColumnName}] >= @StartTime AND [{safeTimeColumnName}] <= @EndTime";
                }

                using (var command = new SqlCommand(query, connection))
                {
                    if (!string.IsNullOrEmpty(timeColumnName) && startTime.HasValue && endTime.HasValue)
                    {
                        command.Parameters.AddWithValue("@StartTime", startTime.Value);
                        command.Parameters.AddWithValue("@EndTime", endTime.Value);
                    }

                    return Convert.ToInt64(command.ExecuteScalar());
                }
            }
        }

        private DataTable GetDataFromSqlServer(string connectionString, string tableName,
                                              List<string> columns, string timeColumnName = null,
                                              string timeColumnDataType = null, DateTime? startTime = null,
                                              DateTime? endTime = null)
        {
            var dataTable = new DataTable();

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // 安全处理表名
                string safeTableName = GetSafeTableName(tableName);

                // 构建列列表
                string columnList = "*";
                if (columns != null && columns.Count > 0)
                {
                    var escapedColumns = columns.Select(col => $"[{col.Trim().Replace("]", "]]")}]");
                    columnList = string.Join(", ", escapedColumns);
                }

                // 构建查询
                string query = $"SELECT {columnList} FROM {safeTableName}";

                // 添加时间过滤
                if (!string.IsNullOrEmpty(timeColumnName) && startTime.HasValue && endTime.HasValue)
                {
                    string safeTimeColumnName = GetSafeColumnName(timeColumnName);

                    if (!string.IsNullOrEmpty(timeColumnDataType))
                    {
                        if (timeColumnDataType.StartsWith("BIGINT") || timeColumnDataType.StartsWith("INT"))
                        {
                            long startTimeValue = long.Parse(startTime.Value.ToString("yyyyMMddHHmm"));
                            long endTimeValue = long.Parse(endTime.Value.ToString("yyyyMMddHHmm"));

                            query += $" WHERE [{safeTimeColumnName}] >= @StartTime AND [{safeTimeColumnName}] <= @EndTime";

                            using (var command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@StartTime", startTimeValue);
                                command.Parameters.AddWithValue("@EndTime", endTimeValue);

                                using (var adapter = new SqlDataAdapter(command))
                                {
                                    adapter.Fill(dataTable);
                                }
                            }
                            return dataTable;
                        }
                    }

                    query += $" WHERE [{safeTimeColumnName}] >= @StartTime AND [{safeTimeColumnName}] <= @EndTime";
                }

                using (var command = new SqlCommand(query, connection))
                {
                    if (!string.IsNullOrEmpty(timeColumnName) && startTime.HasValue && endTime.HasValue)
                    {
                        command.Parameters.AddWithValue("@StartTime", startTime.Value);
                        command.Parameters.AddWithValue("@EndTime", endTime.Value);
                    }

                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
            }

            return dataTable;
        }

        /// <summary>获取表的第一列名（用于 OFFSET/FETCH 的 ORDER BY）</summary>
        private static string GetFirstColumnName(string connectionString, string tableName, List<string> columns)
        {
            if (columns != null && columns.Count > 0)
                return columns[0].Trim().Replace("]", "]]");
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string safeTableName = GetSafeTableName(tableName);
                string query = $"SELECT TOP 0 * FROM {safeTableName}";
                using (var adapter = new SqlDataAdapter(new SqlCommand(query, connection)))
                {
                    var dt = new DataTable();
                    adapter.Fill(dt);
                    if (dt.Columns.Count > 0)
                        return dt.Columns[0].ColumnName.Replace("]", "]]");
                }
            }
            return null;
        }

        /// <summary>分片查询：从 SQL Server 按 OFFSET/FETCH 取一批数据（用于并行导出，避免一次性加载全表）</summary>
        private static DataTable GetChunkFromSqlServer(string connectionString, string tableName,
            List<string> columns, string orderByColumn, int offset, int fetch,
            string timeColumnName, string timeColumnDataType, DateTime? startTime, DateTime? endTime)
        {
            var dataTable = new DataTable();
            if (string.IsNullOrEmpty(orderByColumn)) return dataTable;

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string safeTableName = GetSafeTableName(tableName);
                string columnList = "*";
                if (columns != null && columns.Count > 0)
                {
                    var escaped = columns.Select(c => $"[{c.Trim().Replace("]", "]]")}]");
                    columnList = string.Join(", ", escaped);
                }

                string whereClause = "";
                var paramList = new List<SqlParameter>();

                if (!string.IsNullOrEmpty(timeColumnName) && startTime.HasValue && endTime.HasValue)
                {
                    string safeTimeCol = GetSafeColumnName(timeColumnName);
                    if (!string.IsNullOrEmpty(timeColumnDataType) &&
                        (timeColumnDataType.StartsWith("BIGINT") || timeColumnDataType.StartsWith("INT")))
                    {
                        long startVal = long.Parse(startTime.Value.ToString("yyyyMMddHHmm"));
                        long endVal = long.Parse(endTime.Value.ToString("yyyyMMddHHmm"));
                        whereClause = $" WHERE [{safeTimeCol}] >= @StartTime AND [{safeTimeCol}] <= @EndTime";
                        paramList.Add(new SqlParameter("@StartTime", startVal));
                        paramList.Add(new SqlParameter("@EndTime", endVal));
                    }
                    else
                    {
                        whereClause = $" WHERE [{safeTimeCol}] >= @StartTime AND [{safeTimeCol}] <= @EndTime";
                        paramList.Add(new SqlParameter("@StartTime", startTime.Value));
                        paramList.Add(new SqlParameter("@EndTime", endTime.Value));
                    }
                }

                string orderBy = $"[{orderByColumn}]";
                string query = $"SELECT {columnList} FROM {safeTableName}{whereClause} ORDER BY {orderBy} OFFSET @Offset ROWS FETCH NEXT @Fetch ROWS ONLY";
                paramList.Add(new SqlParameter("@Offset", offset));
                paramList.Add(new SqlParameter("@Fetch", fetch));

                using (var command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddRange(paramList.ToArray());
                    using (var adapter = new SqlDataAdapter(command))
                        adapter.Fill(dataTable);
                }
            }
            return dataTable;
        }

        #endregion

        #region 辅助方法
        private void ValidateParameters(ExcelParams parameters)
        {
            if (string.IsNullOrEmpty(parameters.ConnectionString))
                throw new ArgumentException("连接字符串不能为空");

            if (parameters.SelectedTables == null || parameters.SelectedTables.Count == 0)
                throw new ArgumentException("请选择要导出的表");

            if (string.IsNullOrEmpty(parameters.FilePath))
                throw new ArgumentException("请指定保存路径");

            if (parameters.EnableTimeFilter && parameters.StartTime > parameters.EndTime)
                throw new ArgumentException("开始时间不能晚于结束时间");
        }

        private int CalculateOverallProgress(long rowsExported, long totalRows, int tablesExported, int totalTables)
        {
            if (totalRows == 0 && totalTables == 0) return 0;
            if (totalRows == 0) return (int)((tablesExported * 100.0) / totalTables);

            // 加权计算：70%基于行数，30%基于表数
            double rowsPercentage = (rowsExported * 70.0) / Math.Max(1, totalRows);
            double tablesPercentage = (tablesExported * 30.0) / Math.Max(1, totalTables);

            int progress = (int)(rowsPercentage + tablesPercentage);
            return Math.Max(0, Math.Min(100, progress));
        }

        private string GetValidSheetName(string tableName, int index)
        {
            // Excel工作表名称限制：最大31个字符，不能包含特殊字符
            string sheetName = tableName;

            // 移除或替换非法字符
            var invalidChars = new char[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var ch in invalidChars)
            {
                sheetName = sheetName.Replace(ch, '_');
            }

            // 截断超长名称
            if (sheetName.Length > 31)
            {
                sheetName = sheetName.Substring(0, 31);
            }

            // 确保名称唯一
            if (index > 0)
            {
                string suffix = $"_{index}";
                if (sheetName.Length + suffix.Length > 31)
                {
                    sheetName = sheetName.Substring(0, 31 - suffix.Length);
                }
                sheetName += suffix;
            }

            return sheetName;
        }

        private static string GetSafeTableName(string tableName)
        {
            if (tableName.Contains("."))
            {
                var parts = tableName.Split('.');
                return $"[{parts[0].Trim()}].[{parts[1].Trim()}]";
            }
            return $"[{tableName.Trim()}]";
        }

        private static string GetSafeColumnName(string columnName)
        {
            return columnName.Trim().Replace("]", "]]");
        }
        #endregion

        #region 事件触发方法
        protected virtual void OnProgressChanged(ProgressChangedEventArgs e)
        {
            ProgressChanged?.Invoke(this, e);

            // 如果提供了进度报告器，也通过它报告
            if (e.UserState is ProgressData progressData)
            {
                var parameters = progressData.Parameters as ExcelParams;
                parameters?.ProgressReporter?.ReportProgress(e.ProgressPercentage, progressData.Status);
            }
        }

        protected virtual void OnExportCompleted(ExcelExportResult result)
        {
            ExportCompleted?.Invoke(this, new ExportCompletedEventArgs(result));

            // 如果提供了进度报告器，也记录日志
            if (result != null)
            {
                var logLevel = result.Success ? LogLevel.Success : LogLevel.Error;
                // 这里需要根据实际情况传递参数，或者通过其他方式获取
            }
        }
        #endregion

        #region IDisposable 实现
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 释放托管资源
                    _backgroundWorker?.Dispose();
                    _cancellationTokenSource?.Dispose();
                }

                _disposed = true;
            }
        }

        ~SqlToExcel()
        {
            Dispose(false);
        }
        #endregion

        #region 内部类
        private class TimeColumnInfo
        {
            public string ColumnName { get; set; }
            public string DataType { get; set; }
        }

        private class RowCountInfo
        {
            public long TotalRows { get; set; }
            public Dictionary<string, long> TableRowCounts { get; set; }
        }

        private class ExportStats
        {
            public int TablesExported { get; set; }
            public long RowsExported { get; set; }
        }

        private struct CsvChunkInfo
        {
            public int StartIndex;
            public int RowCount;
            public string FilePath;
        }

        private class ProgressData
        {
            public string Status { get; set; }
            public string CurrentTable { get; set; }
            public int TableProgress { get; set; }
            public int OverallProgress { get; set; }
            public ExcelParams Parameters { get; set; }
        }

        /// <summary>
        /// 导出完成事件参数
        /// </summary>
        public class ExportCompletedEventArgs : EventArgs
        {
            public ExcelExportResult Result { get; }

            public ExportCompletedEventArgs(ExcelExportResult result)
            {
                Result = result;
            }
        }
        #endregion

        #region 临时变量（用于进度计算）
        private long rowsExportedSoFar = 0;
        #endregion
    }
}