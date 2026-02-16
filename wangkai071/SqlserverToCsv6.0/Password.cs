using System;
using System.Windows.Forms;

namespace DatabasePRG
{
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;
            this.AcceptButton = btnOk;  // Enter键触发确定
            this.CancelButton = btnCancel; // Esc键触发取消
        }
        void Password_Load(object sender, EventArgs e)
        {
            //  throw new NotImplementedException();
            // 自动聚焦到密码框
            txtPassword.Focus(); 
        }

        public static bool SlowEquals(string a, string b)
        {
            if (a == null || b == null) return false;
            byte[] aBytes = System.Text.Encoding.UTF8.GetBytes(a);
            byte[] bBytes = System.Text.Encoding.UTF8.GetBytes(b);
            if (aBytes.Length != bBytes.Length) return false;

            int result = 0;
            for (int i = 0; i < aBytes.Length; i++)
                result |= aBytes[i] ^ bBytes[i];
            return result == 0;
        }
        public string EnteredPassword => txtPassword.Text;
        public string Title  {  get => this.Text; set => this.Text = value; }
        public string PromptText  { get => lblPrompt.Text;  set => lblPrompt.Text = value;  }
        public int MaxLength  {  get => txtPassword.MaxLength; set => txtPassword.MaxLength = value;   }
        public Password(string title, string prompt = "请输入管理员密码:") : this()  { this.Title = title; this.PromptText = prompt;  }
        private void txtPassword_TextChanged(object sender, EventArgs e)  {  btnOk.Enabled = !string.IsNullOrWhiteSpace(txtPassword.Text);  }
        private void btnOk_Click(object sender, EventArgs e) { DialogResult = DialogResult.OK;  }
        private void btnCancel_Click(object sender, EventArgs e) {  DialogResult = DialogResult.Cancel;  }
        private void lbl_Click(object sender, EventArgs e) { }
        private void txtPassword_KeyDown(object sender, KeyEventArgs e) {  if (e.KeyCode == Keys.Enter && btnOk.Enabled)  {  btnOk_Click(sender, e);   }  }
    }
}