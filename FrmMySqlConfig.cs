using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QbIntegration.clsHelper;

namespace QbIntegration
{
    public partial class FrmMySqlConfig : Form
    {
        public FrmMySqlConfig()
        {
            InitializeComponent();
        }

        private Boolean CheckValidation()
        {
            Boolean IsValid = false;
            if (txtDatabaseHost.Text.Trim() == "")
            {
                DisplayMessage("Please Enter DatabaseHost Name.", "E");
                IsValid = false;
                txtDatabaseHost.Focus();
            }
            else if (txtDatabaseName.Text.Trim() == "")
            {
                DisplayMessage("Please Enter Database Name.", "E");
                IsValid = false;
                txtDatabaseName.Focus();
            }
            else if (txtUserName.Text.Trim() == "")
            {
                DisplayMessage("Please Enter User Name.", "E");
                IsValid = false;
                txtUserName.Focus();
            }
            else if (txtPassword.Text.Trim() == "")
            {
                DisplayMessage("Please Enter Password.", "E");
                IsValid = false;
                txtPassword.Focus();
            }
            else
                IsValid = true;
            return IsValid;
        }

        private void DisplayMessage(string Text, string Mode)
        {
            switch (Mode)
            {
                case "W":
                    lblErrorMsg.StateCommon.TextColor = Color.FromArgb(16, 6, 244);
                    lblErrorMsg.StateNormal.TextColor = Color.FromArgb(16, 6, 244);
                    lblErrorMsg.Text = Text;
                    break;
                case "I":
                    lblErrorMsg.StateCommon.TextColor = Color.DarkGreen;
                    lblErrorMsg.StateNormal.TextColor = Color.DarkGreen;
                    lblErrorMsg.Text = Text;
                    break;
                case "E":
                    lblErrorMsg.StateCommon.TextColor = Color.DarkRed;
                    lblErrorMsg.StateNormal.TextColor = Color.DarkRed;
                    lblErrorMsg.Text = "Error: " + Text;
                    break;
            }
        }
   

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (Program.mySetting.HostName != "")
                {
                    txtDatabaseHost.Text = Program.mySetting.HostName;
                    txtDatabaseName.Text = Program.mySetting.DatabaseName;
                    txtUserName.Text = Program.mySetting.UserName;
                    txtPassword.Text = Program.mySetting.Password;
                    btnNext.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmMySqlConfig,Function :Form1_Load. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnVerify_Click(object sender, EventArgs e)
        {
            try
            {
                if (CheckValidation())
                {
                    string Connection = "server=" + txtDatabaseHost.Text.Trim() + ";user=" + txtUserName.Text.Trim() + ";Persist Security Info=True;database=" + txtDatabaseName.Text.Trim() + ";port=3306;password=" + txtPassword.Text.Trim() + ";";
                    Boolean IsVerify = clsDBOperation.VerifyConnection_Info(Connection);
                    if (IsVerify)
                    {
                        Program.mySetting.HostName = txtDatabaseHost.Text.Trim();
                        Program.mySetting.DatabaseName = txtDatabaseName.Text.Trim();
                        Program.mySetting.UserName = txtUserName.Text.Trim();
                        Program.mySetting.Password = txtPassword.Text.Trim();
                        Program.mySetting.Connection = Connection;
                        Program.mySetting.Save();
                        DisplayMessage("Database Connection Saved Successfully", "I");
                        btnNext.Enabled = true;
                    }
                    else
                    {
                        DisplayMessage("Database Connection not Done.", "I");
                    }
                }
            }
            catch(Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmMySqlConfig,Function :btnClose_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            try
            {
                ClsCommon.objSqlConfig.Hide();
                ClsCommon.objConfig.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmMySqlConfig,Function :btnNext_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            try
            {
                if (ClsCommon.Mode == "Config")
                {
                    Application.Exit();
                }
                else
                {
                    ClsCommon.objSqlConfig.Hide();
                    ClsCommon.objSync.ToolStripVisibility("SQLConfig");
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmMySqlConfig,Function :btnClose_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        public void Visibility()
        {
            try
            {
                btnNext.Visible = false;
                btnClose.Visible = true;
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmMySqlConfig,Function :Visibility. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_NCHITTEST)
                m.Result = (IntPtr)(HT_CAPTION);
        }

        private const int WM_NCHITTEST = 0x84;
        private const int HT_CLIENT = 0x1;
        private const int HT_CAPTION = 0x2;
    }
}
