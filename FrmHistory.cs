using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;

using System.Xml;
using System.IO;
using QbIntegration.clsHelper;

namespace QbIntegration
{
    public partial class FrmHistory : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public FrmHistory()
        {
            InitializeComponent();
        }

        private void FillGrid()
        {

            dgHistory.Rows.Clear();


            DataTable dtHistory = new DataTable();
            string Date = dtViewDate.Value.Year + "-" + dtViewDate.Value.Month.ToString("00") + "-" + dtViewDate.Value.Day.ToString("00");
            //DateTime dt = Convert.ToDateTime(Date);

            dtHistory = clsDBOperation.RetrieveDataforSync("SELECT * FROM history WHERE DATE(CreatedDate)='"+ Date+"'");
            if (dtHistory.Rows.Count != 0)
            {
                int j = 1;
                for (int i = 0; i < dtHistory.Rows.Count; i++)
                {
                    dgHistory.Rows.Add();
                    dgHistory.Rows[i].Cells[0].Value = (j);
                    dgHistory.Rows[i].Cells[1].Value = Convert.ToDateTime(dtHistory.Rows[i]["CreatedDate"]).ToLongTimeString();
                    dgHistory.Rows[i].Cells[2].Value = dtHistory.Rows[i]["TranType"].ToString();
                    dgHistory.Rows[i].Cells[3].Value = dtHistory.Rows[i]["TotalRec"].ToString();
                    dgHistory.Rows[i].Cells[4].Value = dtHistory.Rows[i]["InsertRec"].ToString();
                    dgHistory.Rows[i].Cells[5].Value = dtHistory.Rows[i]["UpdateRec"].ToString();
                    dgHistory.Rows[i].Cells[6].Value = dtHistory.Rows[i]["SkipRec"].ToString();
                    dgHistory.Rows[i].Cells[7].Value = dtHistory.Rows[i]["FailRec"].ToString();
                    j++;

                }
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
                ClsCommon.WriteErrorLogs("Form:FrmHistory,Function :SetCurrentDate. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void DisplayMessage(string Text, string Mode)
        {
            switch (Mode)
            {
                case "I":
                    lblErrorMessage.StateCommon.TextColor = Color.DarkGreen;
                    lblErrorMessage.StateNormal.TextColor = Color.DarkGreen;
                    lblErrorMessage.Text = Text;
                    break;
                case "E":
                    lblErrorMessage.StateCommon.TextColor = Color.DarkRed;
                    lblErrorMessage.StateNormal.TextColor = Color.DarkRed;
                    lblErrorMessage.Text = "Error: " + Text;
                    break;
                case "W":
                    lblErrorMessage.StateCommon.TextColor = Color.DarkBlue;
                    lblErrorMessage.StateNormal.TextColor = Color.DarkBlue;
                    lblErrorMessage.Text = Text;


                    break;
            }
        }

        private void dtViewDate_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                string Date = dtViewDate.Value.Year + "-" + dtViewDate.Value.Month + "-" + dtViewDate.Value.Day;

                FillGrid();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmScheduler,Function :btnSave_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            ClsCommon.ObjHistory.Hide();
            ClsCommon.objSync.ToolStripVisibility("History");
        }

        private void FrmHistory_Load(object sender, EventArgs e)
        {

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