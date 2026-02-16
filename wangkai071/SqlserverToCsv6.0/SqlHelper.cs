using System;
using System.Data.SqlClient;
using System.Collections.Generic;
public static class SqlHelper//数据库相关类
{
      public class ConnectionTestResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string ConnectionString { get; set; }

        // 静态工厂方法（可选，提升可读性）
        public static ConnectionTestResult SuccessResult(string connStr, string msg = "连接成功")
            => new ConnectionTestResult { Success = true, Message = msg, ConnectionString = connStr };

        public static ConnectionTestResult FailureResult(string msg)
            => new ConnectionTestResult { Success = false, Message = msg, ConnectionString = string.Empty };
    }
      public static ConnectionTestResult TestConnection(string serverName, bool useWindowsAuth, string username = null, string password = null)
    {
        if (string.IsNullOrWhiteSpace(serverName))
            return ConnectionTestResult.FailureResult("服务器名称不能为空");

        try
        {
            string connectionString;
            if (useWindowsAuth)
            {
                connectionString = $"Data Source={serverName};Integrated Security=True;Initial Catalog=master;Connection Timeout=10";
            }
            else
            {
                if (string.IsNullOrWhiteSpace(username))
                    return ConnectionTestResult.FailureResult("用户名不能为空");

                connectionString = $"Data Source={serverName};User ID={username};Password={password};Initial Catalog=master;Connection Timeout=10";
            }

            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
            }

            return ConnectionTestResult.SuccessResult(connectionString, "数据库连接成功！");
        }
        catch (SqlException sqlEx)
        {
            return ConnectionTestResult.FailureResult($"SQL错误({sqlEx.Number}): {sqlEx.Message}");
        }
        catch (Exception ex)
        {
            return ConnectionTestResult.FailureResult($"连接失败: {ex.Message}");
        }
        

    }
      public static (List<string> Databases, string Error) GetDatabaseList(string connectionString)
    {
        try
        {
            List<string> databases = new List<string>();

            using (SqlConnection masterConn = new SqlConnection(connectionString))
            {
                masterConn.Open();

                string sql = @"SELECT name FROM sys.databases 
                                 WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')
                                 AND state = 0  -- 仅显示在线数据库
                                 ORDER BY name";

                using (SqlCommand cmd = new SqlCommand(sql, masterConn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        databases.Add(reader["name"].ToString());
                    }
                }
            }

            return (databases, null);
        }
        catch (Exception ex)
        {
            return (null, $"获取数据库列表失败: {ex.Message}");
        }
    }
      public static (List<string> Tables, string Error) GetTableList(string connectionString, string databaseName)
    {
        try
        {
            List<string> tables = new List<string>();

            // 确保使用正确的数据库连接字符串
            string dbConnectionString = connectionString.Replace("Initial Catalog=master",
                $"Initial Catalog={databaseName}");

            using (SqlConnection conn = new SqlConnection(dbConnectionString))
            {
                conn.Open();

                string sql = @"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES 
                                 WHERE TABLE_TYPE = 'BASE TABLE' 
                                 ORDER BY TABLE_NAME";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        tables.Add(reader["TABLE_NAME"].ToString());
                    }
                }
            }

            return (tables, null);
        }
        catch (Exception ex)
        {
            return (null, $"获取表格列表失败: {ex.Message}");
        }
    }
      public static string CreateDatabaseConnectionString(string baseConnectionString, string databaseName)
    {
        // 这里假设基础连接字符串使用的是 master 数据库
        return baseConnectionString.Replace("Initial Catalog=master",
            $"Initial Catalog={databaseName}");
    }
      public static (bool Success, string Error) DeleteTable( string connectionString, string databaseName, string tableName)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(databaseName))
                return (false, "数据库名称不能为空");

            if (string.IsNullOrWhiteSpace(tableName))
                return (false, "表名称不能为空");

            string dbConnectionString = GetDatabaseConnectionString(connectionString, databaseName);

            using (SqlConnection conn = new SqlConnection(dbConnectionString))
            {
                conn.Open();

                // ✅ 修正：使用 @$"..." 语法（C# 7.3 兼容）
                string sql = $@"
IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
           WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = N'{tableName}')
BEGIN
    DROP TABLE [{tableName}];
    SELECT 1 AS Result;
END
ELSE
BEGIN
    SELECT 0 AS Result;
END";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    object result = cmd.ExecuteScalar();
                    bool tableExists = result != null && Convert.ToInt32(result) == 1;

                    return tableExists
                        ? (true, null)
                        : (false, $"表 '{tableName}' 不存在于数据库 '{databaseName}' 中");
                }
            }
        }
        catch (SqlException sqlEx)
        {
            // ✅ 修正：改用传统 switch 语句（C# 7.3 兼容）
            string errorMessage;
            switch (sqlEx.Number)
            {
                case 3701: // 对象不存在
                    errorMessage = "表不存在或已被删除";
                    break;
                case 2714: // 对象已存在（外键约束等）
                    errorMessage = "表中存在外键约束，请先删除相关约束";
                    break;
                case 1913: // 索引已存在（此处用于表名无效场景）
                    errorMessage = "表不存在或名称无效";
                    break;
                default:
                    errorMessage = $"SQL错误: {sqlEx.Message}";
                    break;
            }
            return (false, errorMessage);
        }
        catch (Exception ex)
        {
            return (false, $"删除表时发生错误: {ex.Message}");
        }
    }
      private static string GetDatabaseConnectionString(string baseConnectionString, string databaseName)
        {
            // 从基础连接字符串中移除可能已存在的数据库名，然后添加目标数据库
            var builder = new SqlConnectionStringBuilder(baseConnectionString)
            {
                InitialCatalog = databaseName
            };
            return builder.ConnectionString;
        }
      public static List<string> GetTableNamesFromDatabase(string connectionString)
    {
        List<string> tableNames = new List<string>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            try
            {
                // 使用配置的超时设置
                connection.Open();

                // 获取用户表（排除系统表）
                string query = @"
                        SELECT TABLE_SCHEMA + '.' + TABLE_NAME as FullTableName
                        FROM INFORMATION_SCHEMA.TABLES 
                        WHERE TABLE_TYPE = 'BASE TABLE'
                        ORDER BY TABLE_SCHEMA, TABLE_NAME";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // 设置命令超时
                    command.CommandTimeout = 30;

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tableNames.Add(reader["FullTableName"].ToString());
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                if (sqlEx.Number == -2) // 超时错误
                {
                    throw new Exception("连接数据库超时，请检查网络连接或增加超时时间。");
                }
                throw new Exception("获取数据库表名失败: " + sqlEx.Message, sqlEx);
            }
            catch (Exception ex)
            {
                throw new Exception("获取数据库表名失败: " + ex.Message, ex);
            }
        }

        return tableNames;
    }//shuaxin_调用
    public static Dictionary<string, string> GetTableColumnsWithTypes(string connectionString, string tableName)
    {
        Dictionary<string, string> columnInfo = new Dictionary<string, string>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            try
            {
                connection.Open();

                // 解析架构和表名
                string schema = "dbo";
                string table = tableName;

                if (tableName.Contains("."))
                {
                    string[] parts = tableName.Split('.');
                    schema = parts[0].Trim().Replace("[", "").Replace("]", "");
                    table = parts[1].Trim().Replace("[", "").Replace("]", "");
                }

                string query = @"
                        SELECT 
                            COLUMN_NAME,
                            DATA_TYPE,
                            CHARACTER_MAXIMUM_LENGTH,
                            NUMERIC_PRECISION,
                            NUMERIC_SCALE
                        FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_SCHEMA = @Schema 
                          AND TABLE_NAME = @TableName
                        ORDER BY ORDINAL_POSITION";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.CommandTimeout = 30;
                    command.Parameters.AddWithValue("@Schema", schema);
                    command.Parameters.AddWithValue("@TableName", table);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string columnName = reader["COLUMN_NAME"].ToString();
                            string dataType = reader["DATA_TYPE"].ToString().ToUpper();

                            // 对于某些数据类型，添加长度信息
                            if (dataType == "VARCHAR" || dataType == "NVARCHAR" || dataType == "CHAR" || dataType == "NCHAR")
                            {
                                object maxLength = reader["CHARACTER_MAXIMUM_LENGTH"];
                                if (maxLength != DBNull.Value && (int)maxLength > 0)
                                {
                                    dataType += $"({maxLength})";
                                }
                                else if (maxLength != DBNull.Value && (int)maxLength == -1)
                                {
                                    dataType += "(MAX)";
                                }
                            }
                            else if (dataType == "DECIMAL" || dataType == "NUMERIC")
                            {
                                object precision = reader["NUMERIC_PRECISION"];
                                object scale = reader["NUMERIC_SCALE"];
                                if (precision != DBNull.Value && scale != DBNull.Value)
                                {
                                    dataType += $"({precision},{scale})";
                                }
                            }

                            columnInfo[columnName] = dataType;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"获取表 '{tableName}' 的列信息失败: " + ex.Message, ex);
            }
        }

        return columnInfo;
    }//clbTableNames_调用

}




