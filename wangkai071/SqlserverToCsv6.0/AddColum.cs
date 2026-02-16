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
    public partial class AddColum : Form
    {
        private TextBox txtColumnName;
        private ComboBox cmbDataType;
        private NumericUpDown numLength;
        private NumericUpDown numPrecision;
        private NumericUpDown numScale;
        private CheckBox chkNullable;
        private CheckBox chkPrimaryKey;
        private TextBox txtDefaultValue;
        private Button btnAdd, btnCancel;
        private string databaseName;
        private string tableName;
        private string connectionString;

        public AddColum(string dbName, string tblName, string connString)
        {
            databaseName = dbName;
            tableName = tblName;
            connectionString = connString.Replace("Initial Catalog=master", $"Initial Catalog={databaseName}");

            InitializeComponent2();
            LoadDataTypes();
        }
        private void InitializeComponent2()
        {
            this.Text = $"添加新列 - {tableName}";
            this.Size = new Size(500, 360);
            this.StartPosition = FormStartPosition.CenterParent;

            // 主布局
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                RowCount = 9,
                ColumnCount = 2
            };

            // 设置列宽
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // 设置行高
            for (int i = 0; i < 8; i++)
            {
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            }
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 列名
            mainLayout.Controls.Add(new Label { Text = "列名:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 0);
            txtColumnName = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(5) };
            mainLayout.Controls.Add(txtColumnName, 1, 0);

            // 数据类型
            mainLayout.Controls.Add(new Label { Text = "数据类型:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 1);
            cmbDataType = new ComboBox { Dock = DockStyle.Fill, Margin = new Padding(5), DropDownStyle = ComboBoxStyle.DropDownList };
            mainLayout.Controls.Add(cmbDataType, 1, 1);

            // 长度/精度
            mainLayout.Controls.Add(new Label { Text = "长度/精度:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 2);
            numLength = new NumericUpDown { Dock = DockStyle.Fill, Margin = new Padding(5), Minimum = 1, Maximum = 8000, Value = 50 };
            mainLayout.Controls.Add(numLength, 1, 2);

            // 小数位数（仅对decimal/numeric类型显示）
            mainLayout.Controls.Add(new Label { Text = "小数位数:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 3);
            numScale = new NumericUpDown { Dock = DockStyle.Fill, Margin = new Padding(5), Minimum = 0, Maximum = 38, Value = 2 };
            mainLayout.Controls.Add(numScale, 1, 3);

            // 允许空值
            mainLayout.Controls.Add(new Label { Text = "允许空值:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 4);
            chkNullable = new CheckBox { Checked = true, Text = "是", AutoSize = true, Margin = new Padding(5) };
            Panel nullablePanel = new Panel { Dock = DockStyle.Fill };
            nullablePanel.Controls.Add(chkNullable);
            mainLayout.Controls.Add(nullablePanel, 1, 4);

            // 是否主键
            mainLayout.Controls.Add(new Label { Text = "是否主键:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 5);
            chkPrimaryKey = new CheckBox { Checked = false, Text = "是", AutoSize = true, Margin = new Padding(5) };
            Panel primaryKeyPanel = new Panel { Dock = DockStyle.Fill };
            primaryKeyPanel.Controls.Add(chkPrimaryKey);
            mainLayout.Controls.Add(primaryKeyPanel, 1, 5);

            // 默认值
            mainLayout.Controls.Add(new Label { Text = "默认值:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, 6);
            txtDefaultValue = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(5) };
            mainLayout.Controls.Add(txtDefaultValue, 1, 6);

            // 按钮面板
            Panel buttonPanel = new Panel { Dock = DockStyle.Fill };
            btnAdd = new Button { Text = "添加", Size = new Size(80, 30), Location = new Point(130, 0) };
            btnAdd.Click += BtnAdd_Click;
            btnCancel = new Button { Text = "取消", Size = new Size(80, 30), Location = new Point(260, 0) };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            buttonPanel.Controls.Add(btnAdd);
            buttonPanel.Controls.Add(btnCancel);

            mainLayout.Controls.Add(buttonPanel, 0, 7);
            mainLayout.SetColumnSpan(buttonPanel, 2);

            this.Controls.Add(mainLayout);

            // 数据类型变化事件
            cmbDataType.SelectedIndexChanged += CmbDataType_SelectedIndexChanged;
        }
        private void LoadDataTypes()
        {
            string[] dataTypes = {
            "int", "bigint", "smallint", "tinyint",
            "varchar", "nvarchar", "char", "nchar",
            "datetime", "date", "time", "datetime2",
            "decimal", "numeric", "float", "real",
            "bit", "uniqueidentifier", "text", "ntext",
            "binary", "varbinary", "image"
        };

            cmbDataType.Items.AddRange(dataTypes);
            cmbDataType.SelectedIndex = 4; // 默认选择varchar
        }
        private void CmbDataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedType = cmbDataType.SelectedItem.ToString().ToLower();

            // 根据数据类型显示/隐藏相关控件
            if (selectedType.Contains("char") || selectedType.Contains("binary") || selectedType == "varbinary")
            {
                numLength.Enabled = true;
                numScale.Enabled = false;
                numLength.Maximum = selectedType.Contains("n") ? 4000 : 8000;
            }
            else if (selectedType == "decimal" || selectedType == "numeric")
            {
                numLength.Enabled = true;
                numScale.Enabled = true;
                numLength.Maximum = 38;
            }
            else
            {
                numLength.Enabled = false;
                numScale.Enabled = false;
            }
        }
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            // 验证输入
            if (string.IsNullOrWhiteSpace(txtColumnName.Text))
            {
                MessageBox.Show("请输入列名！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cmbDataType.SelectedItem == null)
            {
                MessageBox.Show("请选择数据类型！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 检查列名是否已存在
                    string checkSql = $@"
                    SELECT COUNT(*) 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = '{tableName}' 
                    AND COLUMN_NAME = '{txtColumnName.Text}'";

                    using (SqlCommand checkCmd = new SqlCommand(checkSql, conn))
                    {
                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show($"列名 '{txtColumnName.Text}' 已存在！", "错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    // 构建ALTER TABLE语句
                    string sql = BuildAlterTableSql();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();

                        // 如果设置为主键，添加主键约束
                        if (chkPrimaryKey.Checked)
                        {
                            AddPrimaryKeyConstraint(conn);
                        }

                        // 如果设置了默认值，添加默认值约束
                        if (!string.IsNullOrWhiteSpace(txtDefaultValue.Text))
                        {
                            AddDefaultValueConstraint(conn);
                        }

                        MessageBox.Show("列添加成功！", "成功",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加列失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string BuildAlterTableSql()
        {
            string columnName = txtColumnName.Text;
            string dataType = cmbDataType.SelectedItem.ToString();
            string nullable = chkNullable.Checked ? "NULL" : "NOT NULL";

            // 构建完整的数据类型定义
            string typeDefinition = dataType;

            if (dataType.ToLower().Contains("char") || dataType.ToLower().Contains("binary"))
            {
                typeDefinition += $"({numLength.Value})";
            }
            else if (dataType.ToLower() == "decimal" || dataType.ToLower() == "numeric")
            {
                typeDefinition += $"({numLength.Value},{numScale.Value})";
            }

            return $"ALTER TABLE [{tableName}] ADD [{columnName}] {typeDefinition} {nullable}";
        }
        private void AddPrimaryKeyConstraint(SqlConnection conn)
        {
            string columnName = txtColumnName.Text;

            // 首先检查是否已存在主键约束
            string checkPkSql = $@"
            SELECT COUNT(*)
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
            WHERE TABLE_NAME = '{tableName}'
            AND CONSTRAINT_TYPE = 'PRIMARY KEY'";

            using (SqlCommand checkCmd = new SqlCommand(checkPkSql, conn))
            {
                int hasPk = Convert.ToInt32(checkCmd.ExecuteScalar());
                if (hasPk > 0)
                {
                    // 如果已存在主键，需要先删除原来的主键约束
                    string dropPkSql = $@"
                    DECLARE @constraintName NVARCHAR(200)
                    SELECT @constraintName = CONSTRAINT_NAME 
                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
                    WHERE TABLE_NAME = '{tableName}'
                    AND CONSTRAINT_TYPE = 'PRIMARY KEY'
                    
                    IF @constraintName IS NOT NULL
                    BEGIN
                        EXEC('ALTER TABLE [{tableName}] DROP CONSTRAINT [' + @constraintName + ']')
                    END";

                    using (SqlCommand dropCmd = new SqlCommand(dropPkSql, conn))
                    {
                        dropCmd.ExecuteNonQuery();
                    }
                }
            }

            // 添加新的主键约束
            string addPkSql = $"ALTER TABLE [{tableName}] ADD CONSTRAINT PK_{tableName}_{columnName} PRIMARY KEY ([{columnName}])";

            using (SqlCommand cmd = new SqlCommand(addPkSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }
        private void AddDefaultValueConstraint(SqlConnection conn)
        {
            string columnName = txtColumnName.Text;
            string defaultValue = txtDefaultValue.Text;
            string dataType = cmbDataType.SelectedItem.ToString().ToLower();

            // 根据数据类型格式化默认值
            string formattedDefault = defaultValue;
            if (dataType.Contains("char") || dataType.Contains("text") || dataType.Contains("date"))
            {
                formattedDefault = $"'{defaultValue}'";
            }
            else if (dataType == "bit")
            {
                formattedDefault = (defaultValue == "1" || defaultValue.ToLower() == "true") ? "1" : "0";
            }

            string addDefaultSql = $"ALTER TABLE [{tableName}] ADD CONSTRAINT DF_{tableName}_{columnName} DEFAULT ({formattedDefault}) FOR [{columnName}]";

            using (SqlCommand cmd = new SqlCommand(addDefaultSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }
        private void AddColumnForm_Load(object sender, EventArgs e)  {  }
    }
}
