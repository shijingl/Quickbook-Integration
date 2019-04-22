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
using System.IO;

namespace QbIntegration
{
    public partial class FrmAppsLog : Form
    {
        public FrmAppsLog()
        {
            InitializeComponent();
        }

        private void FrmAppsLog_Load(object sender, EventArgs e)
        {
            try
            {
                SetCurrentDate(DateTime.Now);
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmAppsLog,Function :FrmAppsLog_Load. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
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

        public void SetCurrentDate(DateTime dt)
        {
            try
            {
                dtViewDate.Value = dt;
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmAppsLog,Function :SetCurrentDate. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void dtViewDate_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                string Date = dtViewDate.Value.Year + "-" + dtViewDate.Value.Month + "-" + dtViewDate.Value.Day;
                txtLog.Text = "";
                if (File.Exists(Application.StartupPath + @"\Logs\" + Date + "-Log.txt") == true)
                {
                    txtLog.LoadFile(Application.StartupPath + @"\Logs\" + Date + "-Log.txt", RichTextBoxStreamType.PlainText);
                }
                if (txtLog.Text.Trim().Length == 0)
                    txtLog.Text = "Application Has no Log for selected Date.";
                txtLog.ScrollToCaret();
                txtLog.Focus();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmAppsLog,Function :dtViewDate_ValueChanged. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
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
