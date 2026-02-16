using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DatabasePRG
{
    public partial class GantiColumns : Form
    {
        //public ModifyColumnsForm()
        //{
        //    InitializeComponent();
        //}
        private DataGridView dgvColumns1;
        private Button btnSave, btnCancel;
        private string databaseName;
        private string tableName;
        private string connectionString;
        private DataTable originalColumns;
        private List<string> originalColumnNames = new List<string>();
        private Dictionary<string, string> columnTypeMapping = new Dictionary<string, string>();

        public GantiColumns(string dbName, string tblName, string connString)
        {
            databaseName = dbName;
            tableName = tblName;
            connectionString = connString.Replace("Initial Catalog=master", $"Initial Catalog={databaseName}");

            InitializeComponent1();
            LoadColumnInfo();
        }

        private void InitializeComponent1()
        {
            this.Text = $"修改列结构 - {tableName}";
            this.Size = new Size(900, 550);
            this.StartPosition = FormStartPosition.CenterParent;

            // 主布局
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 2
            };

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            // 数据网格
            dgvColumns1 = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,  // 禁止删除行，避免复杂处理
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AllowUserToOrderColumns = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };

            SetupDataGridViewColumns();

            // 按钮面板
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            btnSave = new Button
            {
                Text = "保存修改",
                Size = new Size(100, 35),
                Location = new Point(300, 5)
            };
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Size = new Size(100, 35),
                Location = new Point(410, 5)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            buttonPanel.Controls.Add(btnSave);
            buttonPanel.Controls.Add(btnCancel);

            mainLayout.Controls.Add(dgvColumns1, 0, 0);
            mainLayout.Controls.Add(buttonPanel, 0, 1);

            this.Controls.Add(mainLayout);
        }

        private void SetupDataGridViewColumns()
        {
            // 清除现有列
            dgvColumns1.Columns.Clear();

            // 列名
            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn
            {
                HeaderText = "列名",
                Name = "ColumnName",
                Width = 150,
                ReadOnly = false
            };

            // 数据类型
            DataGridViewComboBoxColumn colType = new DataGridViewComboBoxColumn
            {
                HeaderText = "数据类型",
                Name = "DataType",
                Width = 150,
                FlatStyle = FlatStyle.Flat
            };

            // 填充数据类型选项
            string[] dataTypes = {
            "int", "bigint", "smallint", "tinyint",
            "varchar", "nvarchar", "char", "nchar",
            "datetime", "date", "time", "datetime2",
            "decimal", "numeric", "float", "real",
            "bit", "uniqueidentifier", "text", "ntext",
            "binary", "varbinary", "image"
        };
            colType.Items.AddRange(dataTypes);

            // 长度/精度
            DataGridViewTextBoxColumn colLength = new DataGridViewTextBoxColumn
            {
                HeaderText = "长度/精度",
                Name = "Length",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
            };

            // 小数位数
            DataGridViewTextBoxColumn colScale = new DataGridViewTextBoxColumn
            {
                HeaderText = "小数位数",
                Name = "Scale",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
            };

            // 允许空值
            DataGridViewCheckBoxColumn colNullable = new DataGridViewCheckBoxColumn
            {
                HeaderText = "允许空值",
                Name = "IsNullable",
                Width = 80,
                TrueValue = true,
                FalseValue = false
            };

            // 是否主键
            DataGridViewCheckBoxColumn colPrimary = new DataGridViewCheckBoxColumn
            {
                HeaderText = "主键",
                Name = "IsPrimaryKey",
                Width = 60,
                TrueValue = true,
                FalseValue = false
            };

            // 默认值
            DataGridViewTextBoxColumn colDefault = new DataGridViewTextBoxColumn
            {
                HeaderText = "默认值",
                Name = "DefaultValue",
                Width = 150
            };

            // 原始列名（隐藏列）
            DataGridViewTextBoxColumn colOriginalName = new DataGridViewTextBoxColumn
            {
                HeaderText = "原始列名",
                Name = "OriginalColumnName",
                Visible = false
            };

            // 原始数据类型（隐藏列）
            DataGridViewTextBoxColumn colOriginalType = new DataGridViewTextBoxColumn
            {
                HeaderText = "原始数据类型",
                Name = "OriginalDataType",
                Visible = false
            };

            // 添加列到DataGridView
            dgvColumns1.Columns.Add(colOriginalName);
            dgvColumns1.Columns.Add(colOriginalType);
            dgvColumns1.Columns.Add(colName);
            dgvColumns1.Columns.Add(colType);
            dgvColumns1.Columns.Add(colLength);
            dgvColumns1.Columns.Add(colScale);
            dgvColumns1.Columns.Add(colNullable);
            dgvColumns1.Columns.Add(colPrimary);
            dgvColumns1.Columns.Add(colDefault);
        }

        private void LoadColumnInfo()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string sql = $@"
                SELECT 
                    c.COLUMN_NAME,
                    c.DATA_TYPE,
                    c.CHARACTER_MAXIMUM_LENGTH,
                    c.NUMERIC_PRECISION,
                    c.NUMERIC_SCALE,
                    c.IS_NULLABLE,
                    c.COLUMN_DEFAULT,
                    CASE WHEN k.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PRIMARY_KEY
                FROM INFORMATION_SCHEMA.COLUMNS c
                LEFT JOIN (
                    SELECT ku.TABLE_NAME, ku.COLUMN_NAME
                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                    INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                        ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                    WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
                        AND tc.TABLE_NAME = '{tableName}'
                ) k ON c.TABLE_NAME = k.TABLE_NAME AND c.COLUMN_NAME = k.COLUMN_NAME
                WHERE c.TABLE_NAME = '{tableName}'
                ORDER BY c.ORDINAL_POSITION";

                    originalColumns = new DataTable();
                    using (SqlDataAdapter adapter = new SqlDataAdapter(sql, conn))
                    {
                        adapter.Fill(originalColumns);
                    }

                    // 清空现有行
                    dgvColumns1.Rows.Clear();
                    originalColumnNames.Clear();
                    columnTypeMapping.Clear();

                    // 填充DataGridView
                    foreach (DataRow row in originalColumns.Rows)
                    {
                        string columnName = row["COLUMN_NAME"].ToString();
                        string dataType = row["DATA_TYPE"].ToString();

                        originalColumnNames.Add(columnName);
                        columnTypeMapping[columnName] = dataType;

                        int rowIndex = dgvColumns1.Rows.Add();
                        DataGridViewRow dgvRow = dgvColumns1.Rows[rowIndex];

                        // 保存原始值
                        dgvRow.Cells["OriginalColumnName"].Value = columnName;
                        dgvRow.Cells["OriginalDataType"].Value = dataType;

                        // 设置显示值
                        dgvRow.Cells["ColumnName"].Value = columnName;
                        dgvRow.Cells["DataType"].Value = dataType;

                        // 处理长度/精度
                        if (row["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value)
                        {
                            int length = Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"]);
                            if (length > 0 && length != -1)  // -1 表示 MAX
                            {
                                dgvRow.Cells["Length"].Value = length.ToString();
                            }
                            else if (length == -1)
                            {
                                dgvRow.Cells["Length"].Value = "MAX";
                            }
                        }
                        else if (row["NUMERIC_PRECISION"] != DBNull.Value)
                        {
                            dgvRow.Cells["Length"].Value = row["NUMERIC_PRECISION"].ToString();
                            if (row["NUMERIC_SCALE"] != DBNull.Value)
                            {
                                dgvRow.Cells["Scale"].Value = row["NUMERIC_SCALE"].ToString();
                            }
                        }

                        // 允许空值
                        dgvRow.Cells["IsNullable"].Value = row["IS_NULLABLE"].ToString() == "YES";

                        // 是否主键
                        dgvRow.Cells["IsPrimaryKey"].Value = Convert.ToBoolean(row["IS_PRIMARY_KEY"]);

                        // 默认值
                        object defaultValue = row["COLUMN_DEFAULT"];
                        if (defaultValue != DBNull.Value)
                        {
                            string defaultStr = defaultValue.ToString();
                            // 去除括号和引号
                            if (defaultStr.StartsWith("(") && defaultStr.EndsWith(")"))
                                defaultStr = defaultStr.Substring(1, defaultStr.Length - 2);
                            if (defaultStr.StartsWith("'") && defaultStr.EndsWith("'"))
                                defaultStr = defaultStr.Substring(1, defaultStr.Length - 2);
                            dgvRow.Cells["DefaultValue"].Value = defaultStr;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载列信息失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 验证输入
                if (!ValidateInput())
                    return;

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 开始事务
                    using (SqlTransaction transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            List<string> sqlCommands = new List<string>();

                            // 处理列名修改
                            foreach (DataGridViewRow dgvRow in dgvColumns1.Rows)
                            {
                                if (dgvRow.Cells["ColumnName"].Value == null)
                                    continue;

                                string originalName = dgvRow.Cells["OriginalColumnName"].Value?.ToString();
                                string newName = dgvRow.Cells["ColumnName"].Value.ToString();

                                if (!string.IsNullOrEmpty(originalName) && originalName != newName)
                                {
                                    // 检查新列名是否已存在
                                    string checkSql = $@"
                                SELECT COUNT(*) 
                                FROM INFORMATION_SCHEMA.COLUMNS 
                                WHERE TABLE_NAME = '{tableName}' 
                                AND COLUMN_NAME = '{newName}'";

                                    using (SqlCommand checkCmd = new SqlCommand(checkSql, conn, transaction))
                                    {
                                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                                        if (count > 0)
                                        {
                                            MessageBox.Show($"列名 '{newName}' 已存在，无法重命名", "错误",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            return;
                                        }
                                    }

                                    // 生成重命名命令
                                    string renameSql = $"EXEC sp_rename '{tableName}.{originalName}', '{newName}', 'COLUMN';";
                                    sqlCommands.Add(renameSql);
                                }
                            }

                            // 处理数据类型和其他属性修改
                            foreach (DataGridViewRow dgvRow in dgvColumns1.Rows)
                            {
                                if (dgvRow.Cells["ColumnName"].Value == null)
                                    continue;

                                string columnName = dgvRow.Cells["ColumnName"].Value.ToString();
                                string originalType = dgvRow.Cells["OriginalDataType"].Value?.ToString();
                                string newType = dgvRow.Cells["DataType"].Value?.ToString();
                                string length = dgvRow.Cells["Length"].Value?.ToString();
                                string scale = dgvRow.Cells["Scale"].Value?.ToString();
                                bool isNullable = Convert.ToBoolean(dgvRow.Cells["IsNullable"].Value ?? true);
                                bool isPrimaryKey = Convert.ToBoolean(dgvRow.Cells["IsPrimaryKey"].Value ?? false);
                                string defaultValue = dgvRow.Cells["DefaultValue"].Value?.ToString();

                                // 检查是否需要修改数据类型
                                if (!string.IsNullOrEmpty(originalType) && !string.IsNullOrEmpty(newType))
                                {
                                    // 构建完整的类型定义
                                    string typeDefinition = BuildTypeDefinition(newType, length, scale);

                                    // 构建ALTER COLUMN语句
                                    string nullConstraint = isNullable ? "NULL" : "NOT NULL";
                                    string alterSql = $"ALTER TABLE [{tableName}] ALTER COLUMN [{columnName}] {typeDefinition} {nullConstraint};";
                                    sqlCommands.Add(alterSql);
                                }

                                // 处理默认值（这里简化处理，实际可能需要删除旧默认值再添加新默认值）
                                if (!string.IsNullOrEmpty(defaultValue))
                                {
                                    // 检查是否需要添加默认值约束
                                    string checkDefaultSql = $@"
                                SELECT COUNT(*) 
                                FROM INFORMATION_SCHEMA.COLUMNS 
                                WHERE TABLE_NAME = '{tableName}' 
                                AND COLUMN_NAME = '{columnName}' 
                                AND COLUMN_DEFAULT IS NOT NULL";

                                    using (SqlCommand checkCmd = new SqlCommand(checkDefaultSql, conn, transaction))
                                    {
                                        int hasDefault = Convert.ToInt32(checkCmd.ExecuteScalar());
                                        if (hasDefault == 0)
                                        {
                                            // 添加默认值约束
                                            string defaultSql = $"ALTER TABLE [{tableName}] ADD CONSTRAINT DF_{tableName}_{columnName} DEFAULT ({defaultValue}) FOR [{columnName}];";
                                            sqlCommands.Add(defaultSql);
                                        }
                                    }
                                }
                            }

                            // 执行所有SQL命令
                            foreach (string sql in sqlCommands)
                            {
                                using (SqlCommand cmd = new SqlCommand(sql, conn, transaction))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            // 提交事务
                            transaction.Commit();

                            MessageBox.Show("列结构修改成功！", "成功",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                            this.DialogResult = DialogResult.OK;
                            this.Close();
                        }
                        catch (Exception ex)
                        {
                            // 回滚事务
                            transaction.Rollback();
                            throw new Exception($"修改失败: {ex.Message}", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存修改失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ValidateInput()
        {
            foreach (DataGridViewRow dgvRow in dgvColumns1.Rows)
            {
                // 检查列名
                if (dgvRow.Cells["ColumnName"].Value == null ||
                    string.IsNullOrWhiteSpace(dgvRow.Cells["ColumnName"].Value.ToString()))
                {
                    MessageBox.Show("列名不能为空！", "验证错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 检查数据类型
                if (dgvRow.Cells["DataType"].Value == null)
                {
                    MessageBox.Show("请选择数据类型！", "验证错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                string dataType = dgvRow.Cells["DataType"].Value.ToString();
                string length = dgvRow.Cells["Length"].Value?.ToString();
                string scale = dgvRow.Cells["Scale"].Value?.ToString();

                // 验证长度和精度
                if (dataType.ToLower().Contains("char") || dataType.ToLower().Contains("binary"))
                {
                    if (!string.IsNullOrEmpty(length) && length.ToUpper() != "MAX")
                    {
                        if (!int.TryParse(length, out int len) || len <= 0)
                        {
                            MessageBox.Show($"列 '{dgvRow.Cells["ColumnName"].Value}' 的长度必须是正整数或'MAX'！", "验证错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }
                }
                else if (dataType.ToLower() == "decimal" || dataType.ToLower() == "numeric")
                {
                    if (!int.TryParse(length, out int precision) || precision <= 0 || precision > 38)
                    {
                        MessageBox.Show($"列 '{dgvRow.Cells["ColumnName"].Value}' 的精度必须在1-38之间！", "验证错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (!int.TryParse(scale, out int s) || s < 0 || s > precision)
                    {
                        MessageBox.Show($"列 '{dgvRow.Cells["ColumnName"].Value}' 的小数位数必须在0到精度之间！", "验证错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            return true;
        }

        private string BuildTypeDefinition(string dataType, string length, string scale)
        {
            string typeDefinition = dataType;

            if (dataType.ToLower().Contains("char") || dataType.ToLower().Contains("binary"))
            {
                if (!string.IsNullOrEmpty(length))
                {
                    typeDefinition += $"({length})";
                }
                else
                {
                    typeDefinition += "(50)";  // 默认长度
                }
            }
            else if (dataType.ToLower() == "decimal" || dataType.ToLower() == "numeric")
            {
                if (!string.IsNullOrEmpty(length) && !string.IsNullOrEmpty(scale))
                {
                    typeDefinition += $"({length},{scale})";
                }
                else
                {
                    typeDefinition += "(18,2)";  // 默认精度
                }
            }

            return typeDefinition;
        }
        private void GantiColumns_Load(object sender, EventArgs e)  { }
    }
}
