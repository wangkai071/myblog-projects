// SqlDatabaseService.cs
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace DatabaseCreator01.Services
{
    public class SqlDatabaseService : IDatabaseService  // 改为public
    {
        private readonly ApplicationLogger _logger;
        public SqlDatabaseService(ApplicationLogger logger = null)
        {
            _logger = logger;
        }
        #region 连接相关
        public ConnectionTestResult TestConnection(string serverName, bool useWindowsAuth, string username, string password)
        {
            var result = new ConnectionTestResult();

            try
            {
                string connectionString;
                if (useWindowsAuth)
                {
                    connectionString = $"Server={serverName};Integrated Security=true;Connection Timeout=3";
                }
                else
                {
                    connectionString = $"Server={serverName};User ID={username};Password={password};Connection Timeout=3";
                }

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    result.Success = true;
                    result.Message = "数据库连接成功！";
                    result.ConnectionString = connectionString;
                }
            }
            catch (SqlException sqlEx)
            {
                result.Success = false;
                result.Message = $"SQL错误: {sqlEx.Message}";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"连接失败: {ex.Message}";
            }

            return result;
        }

        public DatabaseListResult GetDatabaseList(string connectionString)
        {
            var result = new DatabaseListResult();

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string sql = @"
                        SELECT name 
                        FROM sys.databases 
                        WHERE database_id > 4 
                        AND state = 0 
                        ORDER BY name";

                    using (var cmd = new SqlCommand(sql, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Databases.Add(reader["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Error = $"获取数据库列表失败: {ex.Message}";
                _logger?.Log(result.Error);
            }

            return result;
        }
        #endregion
        #region 数据库操作
        public OperationResult CreateDatabase(string connectionString, string databaseName)
        {
            var result = new OperationResult();

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string sql = $@"
                        IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{databaseName}')
                        BEGIN
                            CREATE DATABASE [{databaseName}];
                        END";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"数据库 '{databaseName}' 创建成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"创建数据库失败: {ex.Message}";
            }

            return result;
        }

        public OperationResult DeleteDatabase(string connectionString, string databaseName)
        {
            var result = new OperationResult();

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string sql = $@"
                        ALTER DATABASE [{databaseName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                        DROP DATABASE [{databaseName}]";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 60;
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"数据库 '{databaseName}' 删除成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"删除数据库失败: {ex.Message}";
            }

            return result;
        }
        #endregion
        #region 表操作
        public TableListResult GetTableList(string connectionString, string databaseName)
        {
            var result = new TableListResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = @"
                        SELECT TABLE_SCHEMA + '.' + TABLE_NAME as TableName
                        FROM INFORMATION_SCHEMA.TABLES 
                        WHERE TABLE_TYPE = 'BASE TABLE'
                        ORDER BY TABLE_SCHEMA, TABLE_NAME";

                    using (var cmd = new SqlCommand(sql, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Tables.Add(reader["TableName"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Error = $"获取表列表失败: {ex.Message}";
                _logger?.Log(result.Error);
            }

            return result;
        }

        public TableStructureResult GetTableStructure(string connectionString, string databaseName, string tableName)
        {
            var result = new TableStructureResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = @"
                        SELECT 
                            c.name AS ColumnName,
                            t.name AS DataType,
                            c.max_length AS MaxLength,
                            c.is_nullable AS IsNullable,
                            ISNULL((
                                SELECT 1 
                                FROM sys.index_columns ic
                                JOIN sys.indexes i ON ic.object_id = i.object_id 
                                    AND ic.index_id = i.index_id
                                WHERE ic.object_id = c.object_id 
                                    AND ic.column_id = c.column_id
                                    AND i.is_primary_key = 1
                            ), 0) AS IsPrimaryKey
                        FROM sys.columns c
                        JOIN sys.types t ON c.user_type_id = t.user_type_id
                        WHERE c.object_id = OBJECT_ID(@TableName)
                        ORDER BY c.column_id";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@TableName", tableName);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var column = new TableColumn
                                {
                                    Name = reader["ColumnName"].ToString(),
                                    DataType = reader["DataType"].ToString(),
                                    Length = GetLengthDisplay(reader["DataType"].ToString(),
                                        Convert.ToInt32(reader["MaxLength"])),
                                    IsNullable = Convert.ToBoolean(reader["IsNullable"]),
                                    IsPrimaryKey = Convert.ToBoolean(reader["IsPrimaryKey"])
                                };

                                result.Columns.Add(column);

                                if (column.IsPrimaryKey)
                                {
                                    result.PrimaryKeys.Add(column.Name);
                                }
                            }
                        }
                    }

                    result.Message = $"已加载表 '{tableName}' 的结构，共 {result.Columns.Count} 列";
                }
            }
            catch (Exception ex)
            {
                result.Error = $"加载表结构失败: {ex.Message}";
                _logger?.Log(result.Error);
            }

            return result;
        }

        private string GetLengthDisplay(string dataType, int maxLength)
        {
            string lengthDisplay = maxLength.ToString();
            if (maxLength == -1) lengthDisplay = "MAX";

            string dataTypeLower = dataType.ToLower();
            if (!dataTypeLower.Contains("char") &&
                !dataTypeLower.Contains("binary") &&
                dataTypeLower != "varbinary")
            {
                lengthDisplay = "";
            }
            else if (dataTypeLower.Contains("nchar") ||
                     dataTypeLower.Contains("nvarchar"))
            {
                if (maxLength > 0 && maxLength != -1)
                {
                    lengthDisplay = (maxLength / 2).ToString();
                }
            }

            return lengthDisplay;
        }

        public OperationResult CreateTable(string connectionString, string databaseName,
            string tableName, List<TableColumn> columns)
        {
            var result = new OperationResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    // 构建SQL语句
                    string sql = BuildCreateTableSql(tableName, columns);

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"表格 '{tableName}' 在数据库 '{databaseName}' 中创建成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"创建表格失败: {ex.Message}";
            }

            return result;
        }

        private string BuildCreateTableSql(string tableName, List<TableColumn> columns)
        {
            var columnDefs = new List<string>();
            var primaryKeys = new List<string>();

            foreach (var column in columns)
            {
                if (string.IsNullOrWhiteSpace(column.Name)) continue;

                string columnDef = $"[{column.Name}] {column.DataType}";

                // 处理长度
                if ((column.DataType.ToLower().Contains("char") ||
                     column.DataType.ToLower().Contains("varchar")) &&
                    !string.IsNullOrEmpty(column.Length))
                {
                    columnDef += $"({column.Length})";
                }
                else if (column.DataType.ToLower() == "decimal")
                {
                    columnDef += "(18, 2)";
                }

                if (!column.IsNullable)
                {
                    columnDef += " NOT NULL";
                }

                columnDefs.Add(columnDef);

                if (column.IsPrimaryKey)
                {
                    primaryKeys.Add($"[{column.Name}]");
                }
            }

            string sql = $"CREATE TABLE [{tableName}] ({string.Join(", ", columnDefs)}";

            if (primaryKeys.Count > 0)
            {
                sql += $", CONSTRAINT PK_{tableName} PRIMARY KEY ({string.Join(", ", primaryKeys)})";
            }

            sql += ")";
            return sql;
        }

        public OperationResult DeleteTable(string connectionString, string databaseName, string tableName)
        {
            var result = new OperationResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = $"DROP TABLE [{tableName}]";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"表 '{tableName}' 删除成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"删除表失败: {ex.Message}";
            }

            return result;
        }

        public OperationResult AddColumn(string connectionString, string databaseName,
            string tableName, TableColumn column)
        {
            var result = new OperationResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = BuildAddColumnSql(tableName, column);

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"列 '{column.Name}' 添加成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"添加列失败: {ex.Message}";
            }

            return result;
        }

        private string BuildAddColumnSql(string tableName, TableColumn column)
        {
            string columnDef = $"[{column.Name}] {column.DataType}";

            if ((column.DataType.ToLower().Contains("char") ||
                 column.DataType.ToLower().Contains("varchar")) &&
                !string.IsNullOrEmpty(column.Length))
            {
                columnDef += $"({column.Length})";
            }

            if (!column.IsNullable)
            {
                columnDef += " NOT NULL";
            }

            return $"ALTER TABLE [{tableName}] ADD {columnDef}";
        }

        public OperationResult DropColumn(string connectionString, string databaseName,
            string tableName, string columnName)
        {
            var result = new OperationResult();

            try
            {
                string dbConnectionString = BuildDatabaseConnectionString(connectionString, databaseName);

                using (var conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = $"ALTER TABLE [{tableName}] DROP COLUMN [{columnName}]";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        result.Success = true;
                        result.Message = $"列 '{columnName}' 删除成功！";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"删除列失败: {ex.Message}";
            }

            return result;
        }
        #endregion
        #region 数据操作
        public DataTable GetTableData(string connectionString, string tableName,
            List<string> columns = null, string timeColumn = null,
            DateTime? startTime = null, DateTime? endTime = null)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string columnList = "*";
                    if (columns != null && columns.Count > 0)
                    {
                        var escapedColumns = columns.Select(c => $"[{c.Trim().Replace("]", "]]")}]").ToList();
                        columnList = string.Join(", ", escapedColumns);
                    }

                    string query = $"SELECT {columnList} FROM [{tableName}]";

                    // 添加时间过滤
                    if (!string.IsNullOrEmpty(timeColumn) && startTime.HasValue && endTime.HasValue)
                    {
                        query += $" WHERE [{timeColumn}] >= @StartTime AND [{timeColumn}] <= @EndTime";
                    }

                    using (var command = new SqlCommand(query, connection))
                    {
                        command.CommandTimeout = 30;

                        if (!string.IsNullOrEmpty(timeColumn) && startTime.HasValue && endTime.HasValue)
                        {
                            command.Parameters.AddWithValue("@StartTime", startTime.Value);
                            command.Parameters.AddWithValue("@EndTime", endTime.Value);
                        }

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.SelectCommand.CommandTimeout = 30;
                            adapter.Fill(dataTable);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Log($"获取表数据失败: {ex.Message}");
                throw;
            }

            return dataTable;
        }

        public long GetRowCount(string connectionString, string tableName,
            string timeColumn = null, DateTime? startTime = null, DateTime? endTime = null)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = $"SELECT COUNT(*) FROM [{tableName}]";

                    if (!string.IsNullOrEmpty(timeColumn) && startTime.HasValue && endTime.HasValue)
                    {
                        query += $" WHERE [{timeColumn}] >= @StartTime AND [{timeColumn}] <= @EndTime";
                    }

                    using (var command = new SqlCommand(query, connection))
                    {
                        command.CommandTimeout = 30;

                        if (!string.IsNullOrEmpty(timeColumn) && startTime.HasValue && endTime.HasValue)
                        {
                            command.Parameters.AddWithValue("@StartTime", startTime.Value);
                            command.Parameters.AddWithValue("@EndTime", endTime.Value);
                        }

                        return Convert.ToInt64(command.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Log($"获取行数失败: {ex.Message}");
                throw;
            }
        }

        public Dictionary<string, string> GetTableColumnsWithTypes(string connectionString, string tableName)
        {
            Dictionary<string, string> columnInfo = new Dictionary<string, string>();

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = @"
                        SELECT 
                            COLUMN_NAME,
                            DATA_TYPE,
                            CHARACTER_MAXIMUM_LENGTH,
                            NUMERIC_PRECISION,
                            NUMERIC_SCALE
                        FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = @TableName
                        ORDER BY ORDINAL_POSITION";

                    using (var command = new SqlCommand(query, connection))
                    {
                        command.CommandTimeout = 30;
                        command.Parameters.AddWithValue("@TableName", tableName);

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string columnName = reader["COLUMN_NAME"].ToString();
                                string dataType = reader["DATA_TYPE"].ToString().ToUpper();

                                columnInfo[columnName] = dataType;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Log($"获取表列信息失败: {ex.Message}");
                throw;
            }

            return columnInfo;
        }

        public List<string> GetTableNames(string connectionString)
        {
            List<string> tableNames = new List<string>();

            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = @"
                        SELECT TABLE_SCHEMA + '.' + TABLE_NAME as FullTableName
                        FROM INFORMATION_SCHEMA.TABLES 
                        WHERE TABLE_TYPE = 'BASE TABLE'
                        ORDER BY TABLE_SCHEMA, TABLE_NAME";

                    using (var command = new SqlCommand(query, connection))
                    {
                        command.CommandTimeout = 30;

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                tableNames.Add(reader["FullTableName"].ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Log($"获取表名失败: {ex.Message}");
                throw;
            }

            return tableNames;
        }
        #endregion
        #region 辅助方法
        private string BuildDatabaseConnectionString(string baseConnectionString, string databaseName)
        {
            if (baseConnectionString.Contains("Initial Catalog="))
            {
                return baseConnectionString.Replace(
                    "Initial Catalog=master",
                    $"Initial Catalog={databaseName}");
            }
            else
            {
                var builder = new SqlConnectionStringBuilder(baseConnectionString);
                builder.InitialCatalog = databaseName;
                return builder.ConnectionString;
            }
        }
        #endregion
    }
}