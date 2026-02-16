// IDatabaseService.cs
using System;
using System.Collections.Generic;
using System.Data;

namespace DatabaseCreator01.Services
{
    public interface IDatabaseService
    {
        #region 连接相关
        ConnectionTestResult TestConnection(string serverName, bool useWindowsAuth, string username, string password);
        DatabaseListResult GetDatabaseList(string connectionString);
        #endregion

        #region 数据库操作
        OperationResult CreateDatabase(string connectionString, string databaseName);
        OperationResult DeleteDatabase(string connectionString, string databaseName);
        #endregion

        #region 表操作
        TableListResult GetTableList(string connectionString, string databaseName);
        OperationResult CreateTable(string connectionString, string databaseName, string tableName, List<TableColumn> columns);
        OperationResult DeleteTable(string connectionString, string databaseName, string tableName);

        TableStructureResult GetTableStructure(string connectionString, string databaseName, string tableName);
        OperationResult AddColumn(string connectionString, string databaseName, string tableName, TableColumn column);
        OperationResult DropColumn(string connectionString, string databaseName, string tableName, string columnName);
        #endregion

        #region 数据操作
        DataTable GetTableData(string connectionString, string tableName,
            List<string> columns = null, string timeColumn = null,
            DateTime? startTime = null, DateTime? endTime = null);

        long GetRowCount(string connectionString, string tableName,
            string timeColumn = null, DateTime? startTime = null,
            DateTime? endTime = null);

        Dictionary<string, string> GetTableColumnsWithTypes(string connectionString, string tableName);
        List<string> GetTableNames(string connectionString);
        #endregion
    }

    #region 结果类
    public class ConnectionTestResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string ConnectionString { get; set; }
    }

    public class DatabaseListResult
    {
        public List<string> Databases { get; set; } = new List<string>();
        public string Error { get; set; }
    }

    public class TableListResult
    {
        public List<string> Tables { get; set; } = new List<string>();
        public string Error { get; set; }
    }

    public class OperationResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Message { get; set; }
    }

    public class TableStructureResult
    {
        public List<TableColumn> Columns { get; set; } = new List<TableColumn>();
        public List<string> PrimaryKeys { get; set; } = new List<string>();
        public string Error { get; set; }
        public string Message { get; set; }  // 添加这一行
    }

    public class TableColumn
    {
        public string Name { get; set; }
        public string DataType { get; set; }
        public string Length { get; set; }
        public bool IsNullable { get; set; }
        public bool IsPrimaryKey { get; set; }
    }
    #endregion
}