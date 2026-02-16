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
using System.Configuration;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using DatabasePRG.Export;

namespace DatabasePRG
{
    public partial class DatabaseCreatorForm : Form
    {
        #region //开始变量定义区域
        private string connectionString, currentDatabase, currentTable;  // 私有字符串变量，用于//存储数据库连接字符串//当前选中的数据库名//当前选中的表
        private List<string> primaryKeys = new List<string>();  // 创建一个字符串列表，用于存储主键列名
        private DataTable originalDataTable;  // 声明一个 DataTable 对象，用于存储原始数据
        private DataGridView dgvDataTable;  // 声明一个 DataGridView 控件，用于显示数据
        private Dictionary<string, List<string>> tableColumns = new Dictionary<string, List<string>>();  // 创建一个字典，键为表名，值为该表的列名列表
        private Dictionary<string, Dictionary<string, string>> columnDataTypes = new Dictionary<string, Dictionary<string, string>>();  // 创建一个嵌套字典，外层键为表名，内层字典的键为列名，值为数据类型
        private System.Windows.Forms.SaveFileDialog saveFileDialog;  // 声明一个保存文件对话框
        private bool isProcessing = false;  // 布尔标志，表示是否正在处理某个操作
        private bool enabled;  // 布尔变量，可能用于控制某些功能是否启用（但此处未指定类型，可能应为控件属性，注意：这里可能是定义了一个与控件属性同名的字段，但通常不建议这样，可能是多余的）
        private readonly ILogger _log = new RiZhi(); //声明_log字段,且Logger类必须实现ILogger接口。  // 声明一个只读的 ILogger 接口实例，并初始化为 RiZhi 类的新实例
        private readonly List<TextBox> _logList = new List<TextBox>(); // 声明日志目标集合，用于存储 TextBox 控件，以便将日志输出到这些文本框
        private SqlToExcel _exporter;  // 用于导出数据到 CSV
        private JinDuTiao _progressReporter;  // 用于报告进度
        private SqlToExcel _excelExporter;  // 导出器实例，用于将选中的表导出为 CSV 文件
        #endregion
        #region  //Form_Load及初始化
        public DatabaseCreatorForm()
        {
            InitializeComponent(); // 初始化窗体上的所有控件（由Visual Studio设计器自动生成）
            InitializeDataGridView(); // 初始化DataGridView控件的设置
            LoadConfiguration(); // 从配置文件加载应用程序设置
            InitializeLogging(); // 初始化日志记录系统
            InitializeExportComponents(); // 初始化导出相关组件
            SetupProgressReporting(); // 设置进度报告系统
        }// 构造函数：当窗体创建时自动执行
        private void DatabaseCreatorForm_Load(object sender, EventArgs e)
        {
            // 设置服务器名称下拉框的选中项为第2项（索引从0开始）
            cmbServerName.SelectedIndex = 1;

            // 设置DataGridView的选择模式为行标题选择
            dgvColumns.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;

            // 允许多行选择
            dgvColumns.MultiSelect = true;

            // 禁止用户添加新行
            dgvColumns.AllowUserToAddRows = false;

            // 订阅数据错误事件（注意：这里与InitializeDataGridView重复）
            dgvColumns.DataError += DgvColumns_DataError;

            // 初始化DataGridView的列定义
            InitializeDataGridViewColumns();
        }  // 窗体加载事件处理程序（当窗体显示时触发）
        private void InitializeDataGridViewColumns()
        {
            dgvColumns.Columns.Clear();// 清除所有现有列

            // 添加"列名"列（TextBox类型）
            dgvColumns.Columns.Add("ColumnName", "列名");
            dgvColumns.Columns[0].Width = 150;

            // 创建"数据类型"列（ComboBox类型）
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.Name = "DataType"; // 列的唯一标识符
            typeColumn.HeaderText = "数据类型"; // 列标题文本
            typeColumn.Width = 100; // 列宽度

            // 向ComboBox添加常见SQL数据类型选项
            typeColumn.Items.AddRange(new string[] {
        "int", "bigint", "smallint", "tinyint", "decimal", "numeric",
        "money", "smallmoney", "float", "real", "date", "datetime",
        "datetime2", "smalldatetime", "char", "varchar", "nchar",
        "nvarchar", "text", "ntext", "binary", "varbinary",
        "image", "bit", "uniqueidentifier", "xml"
    });

            dgvColumns.Columns.Add(typeColumn);// 将数据类型列添加到DataGridView

            // 添加"长度"列（TextBox类型）
            dgvColumns.Columns.Add("Length", "长度");
            dgvColumns.Columns[2].Width = 80;

            // 创建"允许空值"列（CheckBox类型）
            DataGridViewCheckBoxColumn nullableColumn = new DataGridViewCheckBoxColumn();
            nullableColumn.Name = "Nullable"; // 列的唯一标识符
            nullableColumn.HeaderText = "允许空值"; // 列标题文本
            nullableColumn.Width = 80; // 列宽度
            nullableColumn.TrueValue = true; // 选中时的值
            nullableColumn.FalseValue = false; // 未选中时的值
            dgvColumns.Columns.Add(nullableColumn); // 将允许空值列添加到DataGridView

            // 创建"主键"列（CheckBox类型）
            DataGridViewCheckBoxColumn pkColumn = new DataGridViewCheckBoxColumn();
            pkColumn.Name = "PrimaryKey"; // 列的唯一标识符
            pkColumn.HeaderText = "主键"; // 列标题文本
            pkColumn.Width = 60; // 列宽度
            pkColumn.TrueValue = true; // 选中时的值
            pkColumn.FalseValue = false; // 未选中时的值
            dgvColumns.Columns.Add(pkColumn); // 将主键列添加到DataGridView
        }// 初始化DataGridView列的私有方法
        private void InitializeDataGridView()
        {
            // 订阅DataGridView的数据错误事件，当数据验证失败时触发
            dgvColumns.DataError += DgvColumns_DataError;
        }// 初始化DataGridView控件的私有方法
        private void LoadConfiguration()
        {
            // 从AppSettings读取连接超时设置，如果不存在则使用默认值"3"
            string timeout = ConfigurationManager.AppSettings["ConnectionTimeout"] ?? "3";

            // 从AppSettings读取连接字符串，如果不存在则使用默认连接字符串
            txtConnectionString.Text = ConfigurationManager.AppSettings["Conn"] ??
                "Server=localhost;Database=HZY;Integrated Security=true;";

            // 如果连接字符串不为空，则启用"加载表"按钮
            btnLoadTables.Enabled = !string.IsNullOrEmpty(txtConnectionString.Text);
        } // 从配置文件加载配置信息的私有方法
        private void InitializeLogging()
        {
            // 将4个文本框控件添加到日志列表集合中，用于显示日志信息
            _logList.AddRange(new[] { txtLog1, textBox2, textBox3, textBox4 });

            // 使用Logger类记录启动成功的日志信息
            RiZhi.Success("应用程序已启动...", _logList.ToArray());
        } // 初始化日志记录系统的私有方法
        private void SetupProgressReporting()
        {
            // 创建FormProgressReporter实例，传入进度条、标签、日志控件等
            _progressReporter = new JinDuTiao(
                progressBar1, lblProgress, lblPercentage, _log,
                txtLog1, textBox2, textBox3, textBox4);

            // 创建导出器实例，用于将 SQL 数据导出为 CSV
            _exporter = new SqlToExcel(_log);

            // 订阅导出器的进度变化事件
            _exporter.ProgressChanged += Exporter_ProgressChanged;

            // 订阅导出器的完成事件
            _exporter.ExportCompleted += Exporter_ExportCompleted;
        }  // 设置进度报告系统的私有方法
        private void DgvColumns_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // 如果错误类型是ArgumentException（参数异常）
            if (e.Exception is ArgumentException)
            {
                // 记录错误信息到日志
                _log.LogInfo($"DataGridView 数据错误: {e.Exception.Message} (列: {e.ColumnIndex}, 行: {e.RowIndex})",
                    _logList.ToArray());

                // 设置不抛出异常，防止程序崩溃
                e.ThrowException = false;
            }
        }// DataGridView数据错误事件处理程序
        private void Exporter_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 获取状态信息，如果UserState为空则使用默认格式
            string status = e.UserState?.ToString() ?? $"进度: {e.ProgressPercentage}%";

            // 记录进度信息到日志
            _log.LogInfo(status, _logList.ToArray());

            // 更新UI界面上的进度显示
            UpdateProgressUI(e.ProgressPercentage, status);
        } // 导出器进度变化事件处理程序
        private void Exporter_ExportCompleted(object sender, SqlToExcel.ExportCompletedEventArgs e)
        {
            // 检查是否需要在UI线程上执行（跨线程访问控件需要）
            if (this.InvokeRequired)
            {
                // 通过Invoke在UI线程上执行处理结果的方法
                this.Invoke(new Action(() => HandleExportResult(e.Result)));
            }
            else
            {
                // 当前已在UI线程，直接处理结果
                HandleExportResult(e.Result);
            }
        }// 导出器完成事件处理程序
        private void UpdateProgressUI(int percentage, string status)
        {
            // 创建UI更新操作的委托（匿名方法）
            Action updateAction = () =>
            {
                // 设置进度条的值
                progressBar1.Value = percentage;

                // 设置百分比标签文本
                lblPercentage.Text = $"{percentage}%";

                // 设置进度状态标签文本
                lblProgress.Text = status;
            };

            // 检查是否需要跨线程调用（如果当前不是创建控件的线程）
            if (progressBar1.InvokeRequired)
            {
                // 通过Invoke在UI线程上执行更新操作
                progressBar1.Invoke(updateAction);
            }
            else
            {
                // 当前已在UI线程，直接执行更新操作
                updateAction();
            }
        } // 更新UI进度的私有方法
        private void InitializeExportComponents()
        {
            try
            {
                // 创建进度报告器实例
                _progressReporter = new JinDuTiao(
                    progressBar1, lblProgress, lblPercentage, _log,
                    txtLog1, textBox2, textBox3, textBox4);

                // 创建 CSV 导出器实例
                _excelExporter = new SqlToExcel(_log);

                // 只订阅进度变化事件，不订阅完成事件（注意：这里可能与SetupProgressReporting重复）
                _excelExporter.ProgressChanged += Exporter_ProgressChanged;

                // 记录初始化成功的日志
                RiZhi.Success("CSV导出器初始化完成", _logList.ToArray());
            }
            catch (Exception ex)
            {
                // 记录初始化失败的异常信息
                RiZhi.Error($"初始化导出器失败: {ex.Message}", _logList.ToArray());
            }
        }// 初始化导出组件的私有方法
        #endregion
        #region //数据库连接页面
        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            btnTestConnection.Enabled = false;
            try
            {
                string serverName = cmbServerName.Text?.Trim() ?? string.Empty;
                bool useWindowsAuth = chkWindowsAuth.Checked;
                string username = useWindowsAuth ? null : txtUsername.Text.Trim();
                string password = useWindowsAuth ? null : txtPassword.Text;

                var result = SqlHelper.TestConnection(serverName, useWindowsAuth, username, password);

                _log.LogInfo(result.Message, _logList.ToArray());

                if (result.Success)
                {
                    connectionString = result.ConnectionString;
                    _log.LogInfo(connectionString, _logList.ToArray());

                    btnRefreshDatabases.Enabled = true;
                    btnCreateTable.Enabled = true;
                    btnTestConnection.BackColor = Color.FromArgb(225, 240, 255);
                    btnTestConnection.ForeColor = Color.Black;
                    btnTestConnection.Text = "✓ 连接成功";

                    RefreshDatabaseList();
                }
                else
                {
                    btnTestConnection.Text = "连接测试失败";
                    MessageBox.Show(result.Message, "连接测试失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"程序异常: {ex.Message}";
                _log.LogError(errorMsg, _logList.ToArray());
                MessageBox.Show(errorMsg, "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnTestConnection.Enabled = true;
            }
        }//数据库连接页面_测试连接
        private void RefreshDatabaseList()
        {
            try
            {
                cmbDatabases.Items.Clear();
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                //comboBox2.SelectedIndex = 1;

                // 调用 SqlHelper 的方法
                var result = SqlHelper.GetDatabaseList(connectionString);

                if (!string.IsNullOrEmpty(result.Error))
                {
                    throw new Exception(result.Error);
                }

                List<string> foundDatabases = result.Databases;

                // 填充下拉框
                if (foundDatabases != null && foundDatabases.Count > 0)
                {
                    foreach (string db in foundDatabases)
                    {
                        cmbDatabases.Items.Add(db);
                        comboBox1.Items.Add(db);
                        comboBox2.Items.Add(db);
                    }
                    // 输出到日志

                    _log.LogMessage($"发现 {foundDatabases.Count} 个用户数据库:", _logList.ToArray());
                    foreach (string db in foundDatabases)
                    {
                        _log.LogInfo($"• • • {db}", _logList.ToArray());
                    }
                }
                else
                {
                    _log.LogWarning("未找到任何用户数据库（排除系统库）", _logList.ToArray());
                }

                // 更新UI状态
                if (cmbDatabases.Items.Count > 0)
                {
                    cmbDatabases.SelectedIndex = 0;
                    btnDeleteDatabase.Enabled = true;
                    //  MessageBox.Show(cmbDatabases.SelectedItem.ToString());
                    RefreshTableList(cmbDatabases.SelectedItem.ToString());
                }
                else
                {
                    btnDeleteDatabase.Enabled = false;
                }
                if (comboBox1.Items.Count > 0)
                {
                    comboBox1.SelectedIndex = 0;
                    btnDeleteDatabase.Enabled = true;
                    //  MessageBox.Show(cmbDatabases.SelectedItem.ToString());
                    //  RefreshTableList(comboBox1.SelectedItem.ToString());
                }
                else
                {
                    btnDeleteDatabase.Enabled = false;
                }

                _log.LogInfo($"数据库列表刷新完成", _logList.ToArray());
                // comboBox2.SelectedIndex = 1;
            }
            catch (Exception ex)
            {
                _log.LogError($"刷新数据库列表失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"刷新数据库列表失败: {ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//数据库连接页面_测试连接_找数据库
        #endregion
        #region //数据库管理页面
        private void button1_Click(object sender, EventArgs e)
        {
            string dbName = textBox1.Text.Trim();

            if (string.IsNullOrWhiteSpace(dbName))
            {
                MessageBox.Show("请输入数据库名称", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection masterConn = new SqlConnection(connectionString))
                {
                    masterConn.Open();

                    string sql = $@"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{dbName}') BEGIN  CREATE DATABASE [{dbName}];  END";

                    using (SqlCommand cmd = new SqlCommand(sql, masterConn))
                    {
                        cmd.ExecuteNonQuery();
                        _log.LogSuccess($"数据库 '{dbName}' 创建成功！", _logList.ToArray());
                        MessageBox.Show($"数据库 '{dbName}' 创建成功！");
                        RefreshDatabaseList();
                        textBox1.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"创建数据库失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"创建数据库失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//创建数据库
        private void btnRefreshDatabases_Click(object sender, EventArgs e)
        {
            try
            {
                _log.LogInfo("用户点击了【刷新数据库列表】按钮", _logList.ToArray());
                _log.LogInfo("开始刷新数据库列表...", _logList.ToArray()); ;


                if (string.IsNullOrEmpty(connectionString)) // 检查连接是否有效
                {
                    _log.LogWarning("刷新失败：数据库连接未建立", _logList.ToArray());
                    MessageBox.Show("请先建立数据库连接", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                RefreshDatabaseList();

                _log.LogSuccess("数据库列表刷新完成", _logList.ToArray());
            }
            catch (Exception ex)
            {
                _log.LogError($"刷新数据库列表时发生异常: {ex.Message}");
                MessageBox.Show($"刷新失败: {ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//刷新数据库
        private void BtnDeleteDatabase_Click(object sender, EventArgs e)
        {
            if (cmbDatabases.SelectedItem == null) return;
            string dbName = cmbDatabases.SelectedItem.ToString();

            // 第一步：确认操作
            if (MessageBox.Show($"确定要删除数据库 '{dbName}' 吗？\n此操作不可恢复！",
                "高危操作确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;
            using (var pwdForm = new Password()) // 第二步：密码验证（模态对话框）
            {
                if (pwdForm.ShowDialog() != DialogResult.OK) return;
                string correctPassword = ConfigurationManager.AppSettings["AdminPassword"];
                if (string.IsNullOrEmpty(correctPassword))
                {
                    MessageBox.Show("系统未配置管理员密码，请联系管理员！", "配置错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!Password.SlowEquals(pwdForm.EnteredPassword, correctPassword))// 安全比较（防时序攻击）
                {
                    _log.LogWarning($"数据库删除：密码验证失败（用户:{Environment.UserName}）", _logList.ToArray());
                    MessageBox.Show("密码错误！删除操作已取消。", "验证失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }


            try // 第三步：执行删除（密码验证通过）
            {
                using (SqlConnection masterConn = new SqlConnection(connectionString))
                {
                    masterConn.Open();
                    string sql = $@"
                ALTER DATABASE [{dbName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                DROP DATABASE [{dbName}]";

                    using (SqlCommand cmd = new SqlCommand(sql, masterConn))
                    {
                        cmd.CommandTimeout = 60; // 防止大数据库卡住
                        cmd.ExecuteNonQuery();

                        _log.LogInfo($"✓ 数据库 '{dbName}' 已安全删除（操作人:{Environment.UserName}）", _logList.ToArray());
                        RefreshDatabaseList();
                        MessageBox.Show($"数据库 '{dbName}' 删除成功！", "操作完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (SqlException sqlEx) when (sqlEx.Number == 3701) // 数据库不存在
            {
                _log.LogWarning($"数据库 '{dbName}' 不存在，跳过删除", _logList.ToArray());
                RefreshDatabaseList();
            }
            catch (Exception ex)
            {
                string errorMsg = $"删除数据库 '{dbName}' 失败: {ex.Message}";
                _log.LogError(errorMsg, _logList.ToArray());
                MessageBox.Show(errorMsg, "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }// 删除数据库
        private void cmbDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbDatabases.SelectedItem != null)
                {
                    string selectedDatabase = cmbDatabases.SelectedItem.ToString();
                    _log.LogInfo($"用户选择了数据库: {selectedDatabase}", _logList.ToArray());
                    _log.LogInfo($"数据库连接字符串: {connectionString}", _logList.ToArray());
                    RefreshTableList(selectedDatabase);
                    _log.LogInfo($"当前活动数据库已设置为: {selectedDatabase}", _logList.ToArray());
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"切换数据库时发生错误: {ex.Message}", _logList.ToArray());
            }
        }//选择数据库
        #endregion
        #region //表格管理页面
        private void btnDeleteTable_Click_2(object sender, EventArgs e)
        {
            // 1. 检查是否有选中的表
            if (lstTables.SelectedItem == null)
            {
                MessageBox.Show("请先选择要删除的表", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string databaseName = cmbDatabases.SelectedItem?.ToString();
            string tableName = lstTables.SelectedItem.ToString();

            if (string.IsNullOrEmpty(databaseName))
            {
                MessageBox.Show("请先选择数据库", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. 确认操作
            string confirmMsg = $"确定要删除表 '{tableName}' 吗？\n此操作将永久删除表中的所有数据！";
            if (MessageBox.Show(confirmMsg, "高危操作确认",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            // 3. 密码验证
            using (var pwdForm = new Password())
            {
                pwdForm.Text = "删除表确认";

                if (pwdForm.ShowDialog() != DialogResult.OK)
                {
                    _log.LogInfo($"删除表 '{tableName}'：用户取消操作", _logList.ToArray());
                    return;
                }

                // 验证密码
                string correctPassword = ConfigurationManager.AppSettings["AdminPassword"];
                // string correctPassword = GetSecureAdminPassword();
                if (string.IsNullOrEmpty(correctPassword))
                {
                    MessageBox.Show("系统未配置管理员密码，请联系管理员！", "配置错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!Password.SlowEquals(pwdForm.EnteredPassword, correctPassword))
                {
                    //_log.LogInfo("消息  LogWarning("警告 LogError("错误 LogSuccess("成功

                    _log.LogError($"删除表 '{tableName}'：密码验证失败（用户:{Environment.UserName}）", _logList.ToArray());
                    MessageBox.Show("密码错误！删除操作已取消。", "验证失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }

            // 4. 执行删除
            try
            {
                btnDeleteTable.Enabled = false;
                _log.LogInfo($"正在删除表 '{tableName}'...", _logList.ToArray());

                var result = SqlHelper.DeleteTable(
                    connectionString,
                    databaseName,
                    tableName);

                if (result.Success)
                {
                    _log.LogInfo($"✓ 表 '{tableName}' 已成功删除", _logList.ToArray());
                    MessageBox.Show($"表 '{tableName}' 删除成功！", "操作完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 刷新表格列表
                    RefreshTableList(databaseName);
                }
                else
                {
                    _log.LogError($"删除表失败: {result.Error}", _logList.ToArray());
                    MessageBox.Show($"删除表失败: {result.Error}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"删除表 '{tableName}' 时发生异常: {ex.Message}";
                _log.LogError(errorMsg, _logList.ToArray());
                MessageBox.Show(errorMsg, "严重错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnDeleteTable.Enabled = true;
            }
        }// btnDeleteTable_删除表
        private void button2_Click(object sender, EventArgs e)
        {

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("请先选择一个数据库", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (lstTables.SelectedItem == null)
            {
                MessageBox.Show("请先选择要加载的表", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string databaseName = comboBox1.SelectedItem.ToString();
            string tableName = lstTables.SelectedItem.ToString();

            // 检查连接是否有效
            if (string.IsNullOrEmpty(connectionString))
            {
                MessageBox.Show("请先建立数据库连接", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 显示加载提示
                Cursor = Cursors.WaitCursor;
                button2.Enabled = false;

                // 加载表结构到dgvColumns
                LoadTableStructureToGrid(databaseName, tableName);

                // 可选：更新表名文本框
                txtColumnName.Text = tableName;

                _log.LogInfo($"✓ 成功加载表 '{tableName}' 的结构到编辑区", _logList.ToArray());
            }
            catch (Exception ex)
            {
                string errorMsg = $"加载表结构时发生异常: {ex.Message}";
                _log.LogWarning(errorMsg, _logList.ToArray());
                MessageBox.Show(errorMsg, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                button2.Enabled = true;
            }
        } // button2_加载表
        private void BtnModifyColumns_Click(object sender, EventArgs e)
        {
            _log.LogInfo("用户点击了【修改列】按钮", _logList.ToArray());
            // 从 UI 控件获取当前选中的数据库和表
            string databaseName = comboBox1.SelectedItem?.ToString();
            string tableName = lstTables.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(databaseName) || string.IsNullOrEmpty(tableName))
            {
                MessageBox.Show("请先选择数据库和表格", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 更新类成员变量
            currentDatabase = databaseName;
            currentTable = tableName;

            _log.LogInfo($"打开修改列对话框（数据库: {databaseName}, 表: {tableName}）", _logList.ToArray());
            try
            {
                // 打开修改列的对话框
                using (GantiColumns modifyForm = new GantiColumns(currentDatabase, currentTable, connectionString))
                {
                    if (modifyForm.ShowDialog() == DialogResult.OK)
                    {
                        // 重新加载表格结构到网格
                        LoadTableStructureToGrid(currentDatabase, currentTable);
                        //_log.LogInfo("消息  LogWarning("警告 LogError("错误 LogSuccess("成功
                        _log.LogInfo("列结构修改成功", _logList.ToArray());
                        _log.LogInfo($"✓ 表 '{tableName}' 的列结构修改成功", _logList.ToArray());
                        // 刷新相关UI状态 (假设你有这些按钮，否则可以移除)
                        // btnModifyColumns.Enabled = false;
                        // btnSaveData.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"修改列结构失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"修改列结构失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//BtnModifyColumns_修改列
        private void BtnAddColumnDataTab_Click(object sender, EventArgs e)
        {
            _log.LogInfo("用户点击了【添加列(数据页)】按钮", _logList.ToArray());  // 更新日志
            string databaseName = comboBox1.SelectedItem?.ToString();
            string tableName = lstTables.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(databaseName) || string.IsNullOrEmpty(tableName))
            {
                MessageBox.Show("请先选择数据库和表格", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _log.LogInfo($"打开添加列对话框（数据库: {databaseName}, 表: {tableName}）", _logList.ToArray());
            try
            {
                // 更新类成员变量
                currentDatabase = databaseName;
                currentTable = tableName;

                // 使用直接从 UI 获取的值，而不是成员变量
                using (AddColum addColumnForm = new AddColum(databaseName, tableName, connectionString))
                {
                    if (addColumnForm.ShowDialog() == DialogResult.OK)
                    {
                        // 重新加载表格信息
                        LoadTableInfo(databaseName, tableName);

                        // 如果数据已经加载，重新加载数据
                        if (originalDataTable != null)
                        {
                            // 清除当前数据源
                            dgvDataTable.DataSource = null;
                            originalDataTable = null;

                            // 重新加载数据（如果存在这个方法）
                            // BtnLoadData_Click(sender, e);
                        }

                        // 重新加载表格列表
                        RefreshTableList(databaseName);

                        // 重新加载表格结构到编辑区（如果当前选中的是这个表）
                        if (comboBox1.SelectedItem?.ToString() == databaseName &&
                            lstTables.SelectedItem?.ToString() == tableName)
                        {
                            LoadTableStructureToGrid(databaseName, tableName);
                        }//_log.LogInfo("消息  LogWarning("警告 LogError("错误 LogSuccess("成功
                        _log.LogInfo($"✓ 成功向表 '{tableName}' 添加新列", _logList.ToArray());
                        _log.LogInfo("列添加成功", _logList.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"添加列失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"添加列失败: {ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//BtnAddColumnDataTab_添加列
        private void button3_Click(object sender, EventArgs e)
        {

            _log.LogInfo("用户点击了【删除列】按钮", _logList.ToArray());
            _log.LogInfo("开始处理列删除操作...", _logList.ToArray());
            // 1. 检查是否选择了数据库和表格
            string databaseName = comboBox1.SelectedItem?.ToString();
            string tableName = lstTables.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(databaseName) || string.IsNullOrEmpty(tableName))
            {
                MessageBox.Show("请先选择数据库和表格", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. 获取要删除的列名 - 从列定义网格中获取
            if (dgvColumns.SelectedRows.Count == 0 && dgvColumns.SelectedCells.Count == 0)
            {
                MessageBox.Show("请先在列定义网格中选择要删除的列", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string columnName = "";

            // 从选中的行获取列名
            if (dgvColumns.SelectedRows.Count > 0)
            {
                columnName = dgvColumns.SelectedRows[0].Cells[0].Value?.ToString();
            }
            // 或者从选中的单元格获取
            else if (dgvColumns.SelectedCells.Count > 0)
            {
                int rowIndex = dgvColumns.SelectedCells[0].RowIndex;
                if (rowIndex >= 0 && rowIndex < dgvColumns.Rows.Count)
                {
                    columnName = dgvColumns.Rows[rowIndex].Cells[0].Value?.ToString();
                }
            }

            if (string.IsNullOrEmpty(columnName))
            {
                MessageBox.Show("无法获取有效的列名", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 3. 检查是否是主键列
            if (primaryKeys.Contains(columnName))
            {
                MessageBox.Show($"列 '{columnName}' 是主键列，不能直接删除。请先修改主键约束。", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 4. 确认删除
            DialogResult result = MessageBox.Show($"确定要删除列 '{columnName}' 吗？此操作不可恢复！",
                "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                try
                {
                    // 先验证连接字符串
                    if (string.IsNullOrEmpty(connectionString))
                    {
                        MessageBox.Show("数据库连接不可用，请重新连接数据库", "错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 构建连接字符串连接到指定数据库
                    string dbConnectionString = connectionString.Replace("Initial Catalog=master",
                        $"Initial Catalog={databaseName}");

                    // 确保连接字符串格式正确
                    if (!connectionString.Contains("Initial Catalog=master"))
                    {
                        // 如果连接字符串中没有指定master，可能需要其他方式修改
                        var csBuilder = new SqlConnectionStringBuilder(connectionString);
                        csBuilder.InitialCatalog = databaseName;
                        dbConnectionString = csBuilder.ConnectionString;
                    }
                    _log.LogInfo($"准备从表 '{tableName}' 中删除列 '{columnName}'", _logList.ToArray());
                    using (SqlConnection conn = new SqlConnection(dbConnectionString))
                    {
                        conn.Open();

                        string sql = $"ALTER TABLE [{tableName}] DROP COLUMN [{columnName}]";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.CommandTimeout = 30; // 设置超时时间
                            cmd.ExecuteNonQuery();

                            _log.LogInfo($"列 '{columnName}' 删除成功！");

                            // 从网格中移除对应的行
                            RemoveColumnFromGrid(columnName);

                            // 重新加载表格信息
                            LoadTableInfo(databaseName, tableName);

                            // 重新加载表格结构到编辑区
                            LoadTableStructureToGrid(databaseName, tableName);

                            // 刷新表格列表
                            RefreshTableList(databaseName);

                            MessageBox.Show($"列 '{columnName}' 删除成功！", "操作完成",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    _log.LogInfo($"✓ 列 '{columnName}' 从表 '{tableName}' 中删除成功", _logList.ToArray());
                }
                catch (SqlException sqlEx)
                {
                    // 使用传统的 switch 语句处理不同的 SQL 错误
                    string errorMsg;
                    switch (sqlEx.Number)
                    {
                        case 5074:
                            errorMsg = $"列 '{columnName}' 有约束或索引依赖，请先删除依赖项。";
                            break;
                        case 4922:
                            errorMsg = $"列 '{columnName}' 是索引的一部分，请先删除索引。";
                            break;
                        case 1913:
                            errorMsg = $"列 '{columnName}' 是表的唯一列，不能删除。";
                            break;
                        case 3724:
                            errorMsg = $"列 '{columnName}' 是复制的一部分，不能删除。";
                            break;
                        default:
                            errorMsg = $"删除列失败(SQL错误 {sqlEx.Number}): {sqlEx.Message}";
                            break;
                    }

                    _log.LogInfo(errorMsg, _logList.ToArray());
                    MessageBox.Show(errorMsg, "SQL错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    _log.LogError($"删除列失败: {ex.Message}", _logList.ToArray());
                    MessageBox.Show($"删除列失败: {ex.Message}", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }//button3_删除列
        private void RemoveColumnFromGrid(string columnName)//表格管理页面_button3_移除列
        {
            try
            {
                // 查找包含该列名的行
                for (int i = 0; i < dgvColumns.Rows.Count; i++)
                {
                    var row = dgvColumns.Rows[i];
                    if (row.Cells[0].Value?.ToString() == columnName)
                    {
                        dgvColumns.Rows.RemoveAt(i);
                        _log.LogInfo($"从编辑网格中移除了列 '{columnName}' 的定义");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"从网格中移除列定义时出错: {ex.Message}", _logList.ToArray());
            }
        }
        private void LoadTableStructureToGrid(string databaseName, string tableName)
        {
            try
            {
                // 清空现有的列定义和主键列表
                dgvColumns.Rows.Clear();
                primaryKeys.Clear();

                // 构建连接字符串连接到指定数据库
                string dbConnectionString = connectionString.Replace(
                    "Initial Catalog=master",
                    $"Initial Catalog={databaseName}");

                using (SqlConnection conn = new SqlConnection(dbConnectionString))
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

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@TableName", tableName);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string columnName = reader["ColumnName"].ToString();
                                string dataType = reader["DataType"].ToString();
                                int maxLength = Convert.ToInt32(reader["MaxLength"]);
                                bool isNullable = Convert.ToBoolean(reader["IsNullable"]);
                                bool isPrimaryKey = Convert.ToBoolean(reader["IsPrimaryKey"]);

                                if (isPrimaryKey)
                                {
                                    primaryKeys.Add(columnName);
                                }

                                // 处理长度显示
                                string lengthDisplay = maxLength.ToString();
                                if (maxLength == -1) lengthDisplay = "MAX";

                                if (!dataType.ToLower().Contains("char") &&
                                    !dataType.ToLower().Contains("binary") &&
                                    dataType.ToLower() != "varbinary")
                                {
                                    lengthDisplay = "";
                                }
                                else if (dataType.ToLower().Contains("nchar") ||
                                         dataType.ToLower().Contains("nvarchar"))
                                {
                                    if (maxLength > 0 && maxLength != -1)
                                    {
                                        lengthDisplay = (maxLength / 2).ToString();
                                    }
                                }

                                // 添加到dgvColumns - 特别注意ComboBox列的处理
                                int rowIndex = dgvColumns.Rows.Add();
                                DataGridViewRow row = dgvColumns.Rows[rowIndex];

                                // 设置列名
                                row.Cells["ColumnName"].Value = columnName;

                                // 设置数据类型 - 确保值在ComboBox的选项中
                                DataGridViewComboBoxCell dataTypeCell = (DataGridViewComboBoxCell)row.Cells["DataType"];

                                // 检查数据类型是否在选项中，如果不在，先添加到选项中
                                if (!dataTypeCell.Items.Contains(dataType))
                                {
                                    dataTypeCell.Items.Add(dataType);
                                }

                                // 然后设置值
                                dataTypeCell.Value = dataType;

                                // 设置长度
                                row.Cells["Length"].Value = lengthDisplay;

                                // 设置允许空值
                                row.Cells["Nullable"].Value = isNullable;

                                // 设置主键
                                row.Cells["PrimaryKey"].Value = isPrimaryKey;
                            }
                        }
                    }

                    _log.LogInfo($"已加载表 '{tableName}' 的结构，共 {dgvColumns.Rows.Count} 列", _logList.ToArray());
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"加载表结构失败: {ex.Message}";
                _log.LogError(errorMsg, _logList.ToArray());
                MessageBox.Show(errorMsg, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//表格管理页面_公用方法_加载表结构
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //  LogMessage($"数据库连接字符串: {connectionString}");
            if (comboBox1.SelectedItem != null && !string.IsNullOrEmpty(connectionString))
            {
                dgvColumns.Rows.Clear();
                txtColumnName.Clear();

                string databaseName = comboBox1.SelectedItem.ToString();
                currentDatabase = databaseName; // 更新 currentDatabase
                RefreshTableList(databaseName);
            }
        }// 表格管理页面_comboBox1.
        private void LoadTableInfo(string databaseName, string tableName)
        {
            try
            {
                string dbConnectionString = connectionString.Replace("Initial Catalog=master",
                    $"Initial Catalog={databaseName}");

                using (SqlConnection conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    // 获取列信息
                    string sql = $@"
                SELECT 
                    c.COLUMN_NAME,
                    c.DATA_TYPE,
                    c.CHARACTER_MAXIMUM_LENGTH,
                    c.IS_NULLABLE,
                    CASE WHEN k.COLUMN_NAME IS NOT NULL THEN 'YES' ELSE 'NO' END AS IS_PRIMARY_KEY
                FROM INFORMATION_SCHEMA.COLUMNS c
                LEFT JOIN (
                    SELECT ku.TABLE_NAME, ku.COLUMN_NAME
                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                    INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                        ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                    WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
                ) k ON c.TABLE_NAME = k.TABLE_NAME AND c.COLUMN_NAME = k.COLUMN_NAME
                WHERE c.TABLE_NAME = '{tableName}'
                ORDER BY c.ORDINAL_POSITION";

                    StringBuilder infoBuilder = new StringBuilder();
                    infoBuilder.AppendLine($"表格: {tableName}");
                    infoBuilder.AppendLine($"数据库: {databaseName}");
                    infoBuilder.AppendLine();
                    infoBuilder.AppendLine("列结构:");
                    infoBuilder.AppendLine("--------------------------------------------------");

                    //////////primaryKeys.Clear();
                    //////////columnTypes.Clear();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string columnName = reader["COLUMN_NAME"].ToString();
                            string dataType = reader["DATA_TYPE"].ToString();
                            string maxLength = reader["CHARACTER_MAXIMUM_LENGTH"].ToString();
                            string isNullable = reader["IS_NULLABLE"].ToString();
                            string isPrimaryKey = reader["IS_PRIMARY_KEY"].ToString();

                            string typeInfo = dataType;
                            if (!string.IsNullOrEmpty(maxLength) && maxLength != "-1")
                            {
                                typeInfo += $"({maxLength})";
                            }

                            infoBuilder.AppendLine($"{columnName}");
                            infoBuilder.AppendLine($"  类型: {typeInfo}");
                            infoBuilder.AppendLine($"  可空: {isNullable}");
                            infoBuilder.AppendLine($"  主键: {isPrimaryKey}");
                            infoBuilder.AppendLine();

                            if (isPrimaryKey == "YES")
                            {
                                /// primaryKeys.Add(columnName);
                            }

                            //// columnTypes[columnName] = dataType;
                        }
                    }

                    /// txtTableInfo.Text = infoBuilder.ToString();
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"加载表格信息失败: {ex.Message}", _logList.ToArray());
                /// txtTableInfo.Text = $"加载表格信息失败: {ex.Message}";
            }
        }//表格管理页面_公用方法_加载表格信息
        #region 新表格操作
        private void btnCreateTable_Click(object sender, EventArgs e)
        {

            // 修改：使用 comboBox1 而不是 cmbDatabases
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("请先选择一个数据库", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string tableName = txtColumnName.Text.Trim();
            if (string.IsNullOrWhiteSpace(tableName))
            {
                MessageBox.Show("请输入表格名称", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgvColumns.Rows.Count == 0)
            {
                MessageBox.Show("请至少添加一列", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 修改：使用 comboBox1.SelectedItem 获取数据库名
                string dbName = comboBox1.SelectedItem.ToString();

                // 修改：构建连接字符串时，替换数据库名为 comboBox1 中选中的
                string dbConnectionString = connectionString.Replace("Initial Catalog=master", $"Initial Catalog={dbName}");

                if (TableExists(dbName, tableName)) // 检查表是否存在时也使用正确的数据库名
                {
                    if (MessageBox.Show($"表 '{tableName}' 已存在，是否要覆盖？",
                        "确认覆盖", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    {
                        return;
                    }
                }

                using (SqlConnection conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    // 构建创建表格的 SQL
                    string sql = $"CREATE TABLE [{tableName}] (";

                    List<string> columns = new List<string>();
                    List<string> primaryKeys = new List<string>();

                    foreach (DataGridViewRow row in dgvColumns.Rows)
                    {
                        if (row.Cells[0].Value == null || string.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                            continue;

                        string columnName = row.Cells[0].Value.ToString();
                        string dataType = row.Cells[1].Value?.ToString() ?? "nvarchar";
                        string length = row.Cells[2].Value?.ToString() ?? "50";
                        bool isNullable = Convert.ToBoolean(row.Cells[3].Value ?? true);
                        bool isPrimaryKey = Convert.ToBoolean(row.Cells[4].Value ?? false);

                        // 构建列定义
                        string columnDef = $"[{columnName}] {dataType}";

                        if (dataType.ToLower().Contains("char") || dataType.ToLower().Contains("varchar"))
                        {
                            if (int.TryParse(length, out int len) && len > 0)
                            {
                                columnDef += $"({length})";
                            }
                            else
                            {
                                columnDef += "(50)";
                            }
                        }
                        else if (dataType.ToLower() == "decimal")
                        {
                            columnDef += "(18, 2)";
                        }

                        if (!isNullable)
                        {
                            columnDef += " NOT NULL";
                        }

                        columns.Add(columnDef);

                        if (isPrimaryKey)
                        {
                            primaryKeys.Add(columnName);
                        }
                    }

                    sql += string.Join(", ", columns);

                    // 添加主键约束
                    if (primaryKeys.Count > 0)
                    {
                        sql += $", CONSTRAINT PK_{tableName} PRIMARY KEY ({string.Join(", ", primaryKeys)})";
                    }

                    sql += ")";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                        _log.LogInfo($"表格 '{tableName}' 在数据库 '{dbName}' 中创建成功！", _logList.ToArray());

                        // 清空输入
                        txtColumnName.Clear();
                        dgvColumns.Rows.Clear();

                        // 修改：刷新表格列表时使用 comboBox1 中选中的数据库
                        RefreshTableList(dbName);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"创建表格失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"创建表格失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//新表格创建
        private bool TableExists(string databaseName, string tableName)
        {
            try
            {
                string dbConnectionString = connectionString.Replace(
                    "Initial Catalog=master",
                    $"Initial Catalog={databaseName}");

                using (SqlConnection conn = new SqlConnection(dbConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@TableName", tableName);
                        int count = (int)cmd.ExecuteScalar();
                        return count > 0;
                    }
                }
            }
            catch
            {
                return false;
            }
        }//btnCreateTable_Click检查表是否存在时也使用正确的数据库名
        private void btnRemoveColumn_Click(object sender, EventArgs e)
        {
            // 1. 获取选中的行（包括通过选择单元格选中的行）
            var selectedRows = dgvColumns.SelectedRows.Cast<DataGridViewRow>().ToList();

            // 如果通过单元格选择的，获取包含这些单元格的行
            if (selectedRows.Count == 0 && dgvColumns.SelectedCells.Count > 0)
            {
                // 获取所有选中的单元格所在的行（去重）
                var selectedRowsFromCells = dgvColumns.SelectedCells
                    .Cast<DataGridViewCell>()
                    .Select(cell => cell.OwningRow)
                    .Distinct()
                    .ToList();

                selectedRows = selectedRowsFromCells;
            }

            // 2. 检查是否有选中的行
            if (selectedRows.Count == 0)
            {
                MessageBox.Show("请先选择要删除的列定义行。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3. 排序行索引（从大到小），避免删除时索引变化
            var rowsToDelete = selectedRows.OrderByDescending(r => r.Index).ToList();

            // 4. 【关键防护】检测是否全选（行数等于总行数）
            bool isFullSelection = (selectedRows.Count == dgvColumns.Rows.Count);

            // 5. 明确提示用户
            string confirmMsg = isFullSelection
                ? $"⚠️ 警告：您已选中全部 {dgvColumns.Rows.Count} 行！\n确定要删除所有列定义吗？此操作不可恢复！"
                : $"确定要删除选中的 {selectedRows.Count} 行列定义吗？";

            if (MessageBox.Show(confirmMsg, "删除确认",
                MessageBoxButtons.YesNo,
                isFullSelection ? MessageBoxIcon.Warning : MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            // 6. 安全删除（从后往前删除，避免索引变化）
            int deletedCount = 0;
            foreach (var row in rowsToDelete)
            {
                if (!row.IsNewRow) // 避免删除新行（如果允许添加行）
                {
                    dgvColumns.Rows.Remove(row);
                    deletedCount++;
                }
            }

            // 7. 反馈结果
            _log.LogInfo($"成功删除 {deletedCount} 个列定义", _logList.ToArray());
            MessageBox.Show($"已删除 {deletedCount} 个列定义", "操作完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }// 新表删除列
        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            try
            {
                // 记录操作日志
                _log.LogInfo("用户点击了【添加列】按钮", _logList.ToArray());

                // 添加新行
                dgvColumns.Rows.Add("", "int", "50", true, false);

                // 获取新行的索引
                int newRowIndex = dgvColumns.Rows.Count - 1;

                // 记录详细信息
                _log.LogInfo($"已添加新列定义行（索引: {newRowIndex}）", _logList.ToArray());
                _log.LogInfo($"当前表格共有 {dgvColumns.Rows.Count} 个列定义", _logList.ToArray());

                // 可选：自动选中新添加的行
                if (dgvColumns.Rows.Count > 0)
                {
                    dgvColumns.ClearSelection();
                    dgvColumns.Rows[newRowIndex].Selected = true;
                    dgvColumns.FirstDisplayedScrollingRowIndex = newRowIndex;
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"添加列定义行时发生错误: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"添加列失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } // 新表添加列
        #endregion
        #endregion
        #region //数据管理页面
        private void chkEnableTimeFilter_CheckedChanged(object sender, EventArgs e)
        {
            enabled = chkEnableTimeFilter.Checked;
            cmbTimeColumn.Enabled = enabled;
            dtpStartTime.Enabled = enabled;
            dtpEndTime.Enabled = enabled;
            lblStartTime.Enabled = enabled;
            lblEndTime.Enabled = enabled;
        }//启用时间范围过滤
        private void clbTableNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isProcessing || clbTableNames.SelectedIndex < 0)
                return;

            // 立即视觉反馈：仅保留当前项勾选（在设置 isProcessing 前执行）
            for (int i = 0; i < clbTableNames.Items.Count; i++)
            {
                clbTableNames.SetItemChecked(i, i == clbTableNames.SelectedIndex);
            }

            string selectedTable = clbTableNames.SelectedItem.ToString();
            string connectionString = txtConnectionString.Text;

            try
            {
                isProcessing = true;
                clbTableNames.Enabled = false;
                Cursor = Cursors.WaitCursor;
                // 可选：移除 Application.DoEvents() 避免重入风险（非必须，但推荐）
                // Application.DoEvents();     // === 1. 确保数据已加载（缓存或新加载）===
                if (!columnDataTypes.ContainsKey(selectedTable))
                {
                    var columnInfo = SqlHelper.GetTableColumnsWithTypes(connectionString, selectedTable);
                    columnDataTypes[selectedTable] = columnInfo;
                    tableColumns[selectedTable] = columnInfo.Keys.ToList();
                }  // === 2. 【核心修复】统一使用缓存数据更新 UI（无论是否新加载）===
                var currentColumnInfo = columnDataTypes[selectedTable];
                var timeColumns = new List<string>();

                foreach (var col in currentColumnInfo)
                {
                    string dataType = col.Value.ToUpper();
                    string colNameLower = col.Key.ToLower();

                    bool isDateTimeByType = IsDateTimeType(dataType);
                    bool isDateTimeByName = colNameLower.Contains("time") ||
                                            colNameLower.Contains("date") ||
                                            colNameLower.Contains("create") ||
                                            colNameLower.Contains("update") ||
                                            colNameLower.Contains("modify") ||
                                            colNameLower.Contains("record");  // 修复：移除未定义的 'enabled' 变量（原代码笔误）
                    if (isDateTimeByType || (isDateTimeByName && !IsExcludedType(dataType)))
                    {
                        timeColumns.Add($"{col.Key} ({dataType})");
                    }
                }  // === 3. 安全更新下拉框（主线程操作）===
                if (cmbTimeColumn.InvokeRequired)
                {
                    cmbTimeColumn.Invoke(new Action(() =>
                    {
                        UpdateTimeColumnDropdown(timeColumns);
                    }));
                }
                else
                {
                    UpdateTimeColumnDropdown(timeColumns);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载表列信息失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // 异常时清空避免残留旧数据
                if (cmbTimeColumn.InvokeRequired)
                    cmbTimeColumn.Invoke(new Action(() => ClearTimeColumnDropdown()));
                else
                    ClearTimeColumnDropdown();
            }
            finally
            {
                Cursor = Cursors.Default;
                clbTableNames.Enabled = true;
                isProcessing = false;

            }
        }
        private void UpdateTimeColumnDropdown(List<string> items)
        {
            cmbTimeColumn.Items.Clear();
            cmbTimeColumn.Text = string.Empty; // 【关键】强制清空显示文本
            if (items.Count > 0)
            {
                cmbTimeColumn.Items.AddRange(items.ToArray());
                cmbTimeColumn.SelectedIndex = 0; // 自动更新 Text 为第一项
            }
            else
            {
                cmbTimeColumn.SelectedIndex = -1;
                // Text 已在上方清空，此处无需重复设置
            }
        }
        private void ClearTimeColumnDropdown()
        {
            cmbTimeColumn.Items.Clear();
            cmbTimeColumn.SelectedIndex = -1;
            cmbTimeColumn.Text = string.Empty; // 【关键】确保异常时也清空显示
        }
        private bool IsExcludedType(string dataType)
        {
            if (string.IsNullOrEmpty(dataType))
                return false;

            // 排除一些常见的非日期时间类型，尽管列名可能包含时间关键词
            string[] excludedTypes = {
        "INT", "BIGINT", "FLOAT", "DECIMAL", "VARCHAR", "NVARCHAR",
        "CHAR", "NCHAR", "TEXT", "NTEXT", "BIT", "BOOLEAN", "BINARY"
    };

            string baseType = dataType.Split('(')[0].Trim();

            foreach (var type in excludedTypes)
            {
                if (baseType.Equals(type, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
        private bool IsDateTimeType(string dataType)
        {
            if (string.IsNullOrEmpty(dataType))
                return false;

            // 移除括号内的内容（如精度、时区等），只比较基础类型
            string baseType = dataType.Split('(')[0].Trim().ToUpper();

            // 直接检查常见日期时间类型
            // 这里使用StringComparison.OrdinalIgnoreCase进行不区分大小写的比较
            if (baseType.StartsWith("DATE", StringComparison.OrdinalIgnoreCase) ||
                baseType.StartsWith("TIME", StringComparison.OrdinalIgnoreCase) ||
                baseType.Contains("DATETIME") ||
                baseType.Contains("TIMESTAMP") ||
                baseType.StartsWith("YEAR", StringComparison.OrdinalIgnoreCase) ||
                baseType.StartsWith("INTERVAL", StringComparison.OrdinalIgnoreCase) ||
                baseType.Equals("SMALLDATETIME", StringComparison.OrdinalIgnoreCase) ||
                baseType.Equals("DATETIMEOFFSET", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox2.SelectedItem != null)
                {
                    string selectedDatabase = comboBox2.SelectedItem.ToString();

                    // 记录操作日志
                    _log.LogInfo($"用户选择了数据库: {selectedDatabase}", _logList.ToArray());

                    // 构建完整的连接字符串
                    string dbConnectionString = BuildCompleteConnectionString(selectedDatabase);

                    // 更新两个地方
                    connectionString = dbConnectionString; // 更新成员变量
                    txtConnectionString.Text = dbConnectionString; // 更新文本框

                    _log.LogInfo($"数据库连接字符串: {dbConnectionString}", _logList.ToArray());

                    // 刷新表格列表
                    RefreshTableList(selectedDatabase);

                    _log.LogInfo($"当前活动数据库已设置为: {selectedDatabase}", _logList.ToArray());
                    shuaxin();
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"切换数据库时发生错误: {ex.Message}", _logList.ToArray());
            }
        }
        private string BuildCompleteConnectionString(string databaseName)
        {
            try
            {
                // 方法1：如果已有基础连接字符串，使用SqlConnectionStringBuilder
                if (!string.IsNullOrEmpty(connectionString))
                {
                    var builder = new SqlConnectionStringBuilder(connectionString);

                    // 修改数据库名称
                    builder.InitialCatalog = databaseName;

                    // 确保必要的连接属性存在
                    if (string.IsNullOrEmpty(builder.DataSource))
                    {
                        // 如果没有数据源，从UI获取
                        builder.DataSource = cmbServerName.Text?.Trim();
                    }

                    // 设置认证方式
                    if (chkWindowsAuth.Checked)
                    {
                        builder.IntegratedSecurity = true;
                    }
                    else
                    {
                        builder.UserID = txtUsername.Text.Trim();
                        builder.Password = txtPassword.Text;
                        builder.IntegratedSecurity = false;
                    }

                    // 设置合理的超时时间
                    builder.ConnectTimeout = 15;

                    return builder.ConnectionString;
                }

                // 方法2：从头构建连接字符串
                string serverName = cmbServerName.Text?.Trim() ?? "localhost";
                bool useWindowsAuth = chkWindowsAuth.Checked;

                if (useWindowsAuth)
                {
                    return $"Server={serverName};Database={databaseName};Integrated Security=True;Connection Timeout=15;";
                }
                else
                {
                    string username = txtUsername.Text.Trim();
                    string password = txtPassword.Text;
                    return $"Server={serverName};Database={databaseName};User ID={username};Password={password};Connection Timeout=15;";
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"构建连接字符串失败: {ex.Message}", _logList.ToArray());
                // 返回一个基本的连接字符串作为备用
                return $"Server={cmbServerName.Text};Database={databaseName};Integrated Security=True;";
            }
        }  // comboBox2_构建完整的连接字符串
        private void shuaxin()
        {
            try
            {
                string connectionString = txtConnectionString.Text;
                if (string.IsNullOrEmpty(connectionString))
                {
                    MessageBox.Show("请先填写连接字符串", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 显示加载提示
                Cursor = Cursors.WaitCursor;
                //   lblProgress.Text = "正在加载表信息...";
                Application.DoEvents();

                // 获取表名
                List<string> tableNames = SqlHelper.GetTableNamesFromDatabase(connectionString);

                // 更新CheckedListBox
                clbTableNames.BeginUpdate();
                clbTableNames.Items.Clear();
                foreach (var tableName in tableNames)
                {
                    clbTableNames.Items.Add(tableName, false);
                }
                clbTableNames.EndUpdate();

                // 清空列信息
                tableColumns.Clear();
                columnDataTypes.Clear();
                cmbTimeColumn.Items.Clear();

                // 如果表数量很多，调整CheckedListBox大小
                if (tableNames.Count > 10)
                {
                    clbTableNames.Height = Math.Min(tableNames.Count * 20 + 10, 200);
                }

                MessageBox.Show($"成功加载 {tableNames.Count} 个表\n请选择要导出的表", "成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"数据库连接失败：{sqlEx.Message}", "数据库错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载表名失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                //  lblProgress.Text = "就绪";
            }
        }// comboBox2_加载表信息
        #region 导出到 CSV
        private async void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // 检查连接字符串
                if (string.IsNullOrEmpty(txtConnectionString.Text))
                {
                    MessageBox.Show("请填写连接字符串", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查是否选择了表
                var selectedTables = new List<string>();
                foreach (var item in clbTableNames.CheckedItems)
                {
                    selectedTables.Add(item.ToString());
                }

                if (selectedTables.Count == 0)
                {
                    MessageBox.Show("请选择要导出的表", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查时间范围
                if (chkEnableTimeFilter.Checked && dtpStartTime.Value > dtpEndTime.Value)
                {
                    MessageBox.Show("开始时间不能晚于结束时间", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 保存文件对话框
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "CSV文件|*.csv|所有文件|*.*";
                    saveDialog.FileName = $"Export_{DateTime.Now:yyyyMMddHHmmss}.csv";
                    saveDialog.OverwritePrompt = true;

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 1. 准备参数
                        var parameters = new ExcelParams
                        {
                            ConnectionString = txtConnectionString.Text,
                            EnableTimeFilter = chkEnableTimeFilter.Checked,
                            TimeColumnDisplay = cmbTimeColumn.SelectedItem?.ToString(),
                            SelectedTables = selectedTables,
                            FilePath = saveDialog.FileName,
                            StartTime = dtpStartTime.Value,
                            EndTime = dtpEndTime.Value,
                            ProgressReporter = _progressReporter,
                            TableColumns = tableColumns,
                            ColumnDataTypes = columnDataTypes
                        };

                        // 2. 重置UI状态
                        ResetExportUI();

                        // 3. 更新按钮状态
                        btnExport.Enabled = false;
                        btnCancelExport.Visible = true;

                        // 4. 执行导出
                        _log.LogInfo($"开始导出到: {saveDialog.FileName}", _logList.ToArray());
                        _log.LogInfo($"导出表数量: {selectedTables.Count}", _logList.ToArray());

                        // 使用 await 等待导出完成
                        var result = await _excelExporter.ExportAsync(parameters);

                        // 5. 处理结果（这里只会被调用一次）
                        HandleExportResult(result);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"导出失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"导出失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 恢复UI状态
                btnExport.Enabled = true;
                btnCancelExport.Visible = false;
                progressBar1.Visible = false;
            }
        } // 导出按钮点击事件
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            if (disposing)
            {
                // 取消订阅事件并释放导出器
                if (_excelExporter != null)
                {
                    _excelExporter.ProgressChanged -= Exporter_ProgressChanged;
                    _excelExporter.Dispose();
                    _excelExporter = null;
                }

                _progressReporter = null;
            }

            base.Dispose(disposing);
        }
        private void ResetExportUI()
        {
            try
            {
                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new Action(ResetExportUI));
                    return;
                }

                progressBar1.Visible = true;
                progressBar1.Value = 0;
                progressBar1.Style = ProgressBarStyle.Blocks;

                if (lblProgress != null)
                    lblProgress.Text = "准备导出...";

                if (lblPercentage != null)
                    lblPercentage.Text = "0%";

                // 强制UI刷新
                Application.DoEvents();

                _log.LogInfo("UI重置完成", _logList.ToArray());
            }
            catch (Exception ex)
            {
                _log.LogError($"重置UI失败: {ex.Message}", _logList.ToArray());
            }
        }
        private void HandleExportResult(ExcelExportResult result)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => HandleExportResult(result)));
                    return;
                }

                // 隐藏进度条
                progressBar1.Visible = false;

                // 恢复UI状态
                btnExport.Enabled = true;
                btnCancelExport.Visible = false;
                lblProgress.Text = "就绪";
                lblPercentage.Text = "0%";

                if (result == null)
                {
                    MessageBox.Show("导出结果为空", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (result.Success)
                {
                    string message = $"{result.Message}\n";
                    message += $"导出文件: {result.FilePath}\n";
                    message += $"耗时: {result.ElapsedTime:hh\\:mm\\:ss}";

                    _log.LogSuccess(message, _logList.ToArray());

                    var dialogResult = MessageBox.Show(
                        $"{message}\n\n是否打开导出文件所在文件夹？",
                        "导出成功",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (dialogResult == DialogResult.Yes)
                    {
                        try
                        {
                            // 打开文件所在文件夹并选中文件
                            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{result.FilePath}\"");
                        }
                        catch (Exception ex)
                        {
                            _log.LogWarning($"无法打开文件夹: {ex.Message}", _logList.ToArray());
                        }
                    }
                }
                else
                {
                    string errorMessage = result.Message;
                    if (result.Error != null)
                    {
                        errorMessage += $"\n详细信息: {result.Error.Message}";
                    }

                    _log.LogError(errorMessage, _logList.ToArray());
                    MessageBox.Show(errorMessage, "导出失败",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"处理导出结果失败: {ex.Message}", _logList.ToArray());
                MessageBox.Show($"处理导出结果失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnCancelExport_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (_excelExporter != null && _excelExporter.IsExporting)
                {
                    _excelExporter.CancelExport();
                    _progressReporter.LogMessage("正在取消导出...", LogLevel.Warning);
                    _log.LogWarning("用户取消了导出操作", _logList.ToArray());
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"取消导出失败: {ex.Message}", _logList.ToArray());
            }
        }  // 取消导出按钮点击事件
        #endregion
        #endregion
        #region //公用方法
        private void RefreshTableList(string databaseName)
        {

            try
            {
                // 刷新两个列表
                lstTables.Items.Clear();
                var result = SqlHelper.GetTableList(connectionString, databaseName);

                if (!string.IsNullOrEmpty(result.Error))
                {
                    throw new Exception(result.Error);
                }

                List<string> tables = result.Tables;

                // 填充两个列表
                if (tables != null && tables.Count > 0)
                {
                    foreach (string tableName in tables)
                    {

                        lstTables.Items.Add(tableName);
                    }

                    // 设置默认选中项
                    if (lstTables.Items.Count > 0)
                    {
                        lstTables.SelectedIndex = 0;
                    }


                    btnDeleteTable.Enabled = true;
                }
                else
                {
                    btnDeleteTable.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                _log.LogError($"刷新表格列表失败: {ex.Message}", _logList.ToArray());
            }
        }//公用方法_刷新表格列表
        #endregion
        #region //winform控件
        private string GetSecureAdminPassword() { return "19850714"; }// 从安全位置获取密码
        private void btnDeleteTable_Click(object sender, EventArgs e) { }
        private void progressBar1_Click(object sender, EventArgs e) { }
        private void label4_Click(object sender, EventArgs e) { }
        private void panel1_Paint(object sender, PaintEventArgs e) { }
        private void chkWindowsAuth_CheckedChanged(object sender, EventArgs e) { bool enabled = !chkWindowsAuth.Checked; txtUsername.Enabled = enabled; txtPassword.Enabled = enabled; }
        private void txtDatabaseName_TextChanged(object sender, EventArgs e) { }
        private void rightPanel_Paint(object sender, PaintEventArgs e) { }
        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e) { }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e) { }
        private void logGroup_Enter(object sender, EventArgs e) { }
        private void textBox2_TextChanged(object sender, EventArgs e) { }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e) { }
        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e) { }
        private void dgvColumns_CellContentClick(object sender, DataGridViewCellEventArgs e) { }
        private void lblServer_Click(object sender, EventArgs e) { }
        private void cmbServerName_SelectedIndexChanged(object sender, EventArgs e) { }
        private void txtPassword_TextChanged(object sender, EventArgs e) { }
        private void btnCreateDatabase_Click(object sender, EventArgs e) { }
        private void spacer_Paint(object sender, PaintEventArgs e) { }
        private void btnLoadTables_Click(object sender, EventArgs e) { shuaxin(); }
        private void GetTableNamesFromDatabase_TextChanged(object sender, EventArgs e) { }
        private void dataTab_Click(object sender, EventArgs e) { }
        private void lblUser_Click(object sender, EventArgs e) { }
        private void button3_Click_1(object sender, EventArgs e) { }
        private void cmbTimeColumn_SelectedIndexChanged(object sender, EventArgs e) { }
        #endregion

    }
}
//整体架构的优势
//1.单一职责原则
//text
//原设计（一个类做所有事）：
//DatabaseCreatorForm（2000+行）
//├── 连接数据库
//├── 创建表
//├── 删除表
//├── 导出 CSV
//└── UI更新

//新设计（每个类有明确职责）：
//├── DatabaseCreatorForm（UI逻辑，500行）
//├── SqlToExcel（CSV 导出逻辑）
//├── ExcelParams（导出参数封装）
//└── JinDuTiao（UI适配器）
//2. 开闭原则
//对扩展开放：新增导出格式（如PDF）只需创建新类

//对修改封闭：修改导出逻辑不影响主窗体

//3. 依赖倒置原则
//csharp
//// 高层模块不依赖低层模块，都依赖抽象
//public class SqlToExcel
//{
//    // 依赖接口，不依赖具体实现
//    private readonly IProgressReporter _progressReporter;

//    public SqlToExcel(IProgressReporter progressReporter)
//    {
//        _progressReporter = progressReporter;
//    }
//}
//4.可测试性
//csharp
//// 可以创建模拟对象进行测试
//[Test]
//public void Export_ValidParameters_ReturnsSuccess()
//{
//    // 创建模拟进度报告器
//    var mockReporter = new Mock<IProgressReporter>();

//    // 创建导出器
//    var exporter = new SqlToExcel(mockReporter.Object);

//    // 执行测试
//    var result = exporter.Export(parameters);

//    // 验证结果
//    Assert.IsTrue(result.Success);
//}
//5.错误处理集中化
//csharp
//public ExcelExportResult Export(ExcelParams parameters)
//{
//    try
//    {
//        // 所有业务逻辑
//        return ExportInternal(parameters);
//    }
//    catch (SqlException ex)
//    {
//        // 统一处理数据库错误
//        return new ExcelExportResult
//        {
//            Success = false,
//            Message = $"数据库错误: {ex.Message}"
//        };
//    }
//    catch (IOException ex)
//    {
//        // 统一处理文件错误
//        return new ExcelExportResult
//        {
//            Success = false,
//            Message = $"文件错误: {ex.Message}"
//        };
//    }
//}