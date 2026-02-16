// ExcelExportService.cs
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace DatabaseCreator01.Services
{
    public class ExcelExportService
    {
        private readonly ApplicationLogger _logger;

        public ExcelExportService(ApplicationLogger logger)
        {
            _logger = logger;
        }

        public ExportResult ExportToExcel(string connectionString, List<string> tables,
            string filePath, bool enableTimeFilter = false, string timeColumn = null,
            DateTime? startTime = null, DateTime? endTime = null)
        {
            var result = new ExportResult();

            try
            {
                using (var package = new ExcelPackage())
                {
                    int exportedTables = 0;
                    long exportedRows = 0;
                    var databaseService = new SqlDatabaseService(_logger);

                    for (int i = 0; i < tables.Count; i++)
                    {
                        string tableName = tables[i];

                        try
                        {
                            // 获取数据
                            DataTable dataTable = databaseService.GetTableData(
                                connectionString, tableName,
                                timeColumn: enableTimeFilter ? timeColumn : null,
                                startTime: startTime, endTime: endTime);

                            // 创建工作表
                            string sheetName = GetValidSheetName(tableName, i);
                            var worksheet = package.Workbook.Worksheets.Add(sheetName);

                            // 写入列标题
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                worksheet.Cells[1, j + 1].Value = dataTable.Columns[j].ColumnName;
                                worksheet.Cells[1, j + 1].Style.Font.Bold = true;
                            }

                            // 写入数据
                            for (int j = 0; j < dataTable.Rows.Count; j++)
                            {
                                for (int k = 0; k < dataTable.Columns.Count; k++)
                                {
                                    worksheet.Cells[j + 2, k + 1].Value = dataTable.Rows[j][k];
                                }
                            }

                            // 调整列宽
                            if (dataTable.Columns.Count > 0)
                            {
                                int lastRow = dataTable.Rows.Count > 0 ? dataTable.Rows.Count + 1 : 1;
                                worksheet.Cells[1, 1, lastRow, dataTable.Columns.Count].AutoFitColumns();
                            }

                            exportedRows += dataTable.Rows.Count;
                            exportedTables++;

                            _logger?.Log($"表 '{tableName}' 导出完成 ({dataTable.Rows.Count} 条)");
                        }
                        catch (Exception ex)
                        {
                            _logger?.Log($"表 '{tableName}' 导出失败: {ex.Message}");
                        }
                    }

                    // 保存文件
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);

                    result.Success = true;
                    result.ExportedTables = exportedTables;
                    result.ExportedRows = exportedRows;
                    result.FilePath = filePath;
                    result.Message = $"成功导出 {exportedTables} 个表，共 {exportedRows} 条数据";
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"导出失败: {ex.Message}";
                _logger?.Log($"导出失败: {ex.Message}");
            }

            return result;
        }

        private string GetValidSheetName(string tableName, int index)
        {
            string sheetName = tableName;
            if (sheetName.Length > 31)
            {
                sheetName = sheetName.Substring(0, 31);
            }

            sheetName = sheetName.Replace(':', '_')
                                .Replace('\\', '_')
                                .Replace('/', '_')
                                .Replace('?', '_')
                                .Replace('*', '_')
                                .Replace('[', '_')
                                .Replace(']', '_');

            if (index > 0)
            {
                sheetName = $"T{index}_{sheetName}";
                if (sheetName.Length > 31)
                {
                    sheetName = sheetName.Substring(0, 31);
                }
            }

            return sheetName;
        }
    }

    public class ExportResult
    {
        public bool Success { get; set; }
        public int ExportedTables { get; set; }
        public long ExportedRows { get; set; }
        public string FilePath { get; set; }
        public string Message { get; set; }
    }
}