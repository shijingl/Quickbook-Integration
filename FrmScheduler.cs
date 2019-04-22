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
    public partial class FrmScheduler : Form
    {
        DataTable dtScheduler = new DataTable();
        DataTable dt = new DataTable();

        public FrmScheduler()
        {
            InitializeComponent();
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

        private void FrmScheduler_Load(object sender, EventArgs e)
        {
            try
            {

                dt.Columns.Add("Day", typeof(string));
                dt.Columns.Add("StartTime", typeof(string));
                dt.Columns.Add("EndTime", typeof(string));
                dt.Columns.Add("ExecutionTime", typeof(string));
                dt.Columns.Add("Recursive", typeof(string));
                dt.Rows.Add("Monday");
                dt.Rows.Add("Tuesday");
                dt.Rows.Add("Wednesday");
                dt.Rows.Add("Thursday");
                dt.Rows.Add("Friday");
                dt.Rows.Add("Saturday");
                dt.Rows.Add("Sunday");


                // FillDay();
                DataTable dtScheduler = clsDBOperation.RetrieveDataforSync("select * from scheduler");
                if (dtScheduler.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dtScheduler.Rows)
                    {
                        i = dgScheduler.Rows.Count;
                        dgScheduler.Rows.Add();
                        dgScheduler.Rows[i].Cells[0].Value = dr["day"].ToString();
                        dgScheduler.Rows[i].Cells[1].Value = dr["start_time"].ToString();
                        dgScheduler.Rows[i].Cells[2].Value = dr["end_time"].ToString();
                        dgScheduler.Rows[i].Cells[3].Value = dr["execution_time"].ToString();
                        dgScheduler.Rows[i].Cells[4].Value = dr["recursive"].ToString();
                    }
                }
                else
                {
                    FillDay();
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmScheduler,Function :FrmScheduler_Load. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void FillDay()
        {
            try
            {
               

                int i = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    i = dgScheduler.Rows.Count;
                    dgScheduler.Rows.Add();
                    dgScheduler.Rows[i].Cells[0].Value = dr["Day"].ToString();
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmScheduler,Function :FillDay. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                TimeSpan t = TimeSpan.Parse("23:59:59");
                for (int i = 0; i < dgScheduler.Rows.Count; i++)
                {
                    if (dgScheduler.Rows[i].Cells[1].Value != null && dgScheduler.Rows[i].Cells[1].Value.ToString() != "")
                    {
                        if (dgScheduler.Rows[i].Cells[4].Value.ToString() != "")
                        {
                            TimeSpan t1 = TimeSpan.Parse(dgScheduler.Rows[i].Cells[2].Value.ToString());
                            TimeSpan t2 = TimeSpan.Parse(dgScheduler.Rows[i].Cells[3].Value.ToString());
                            if (t1 > t)
                            {
                                DisplayMessage("Enter Valid EndTime", "E");
                                goto Final;
                            }
                            if (t2 > t || t2 > t1)
                            {
                                DisplayMessage("Enter Valid ExecutionTime", "E");
                                goto Final;
                            }
                        }
                        else
                        {
                            DisplayMessage("Please Enter Recursive", "E");
                            goto Final;
                        }
                    }
                }

                //First Delete rows then create it...
                clsDBOperation.UpdateRecords("delete from scheduler");
                if (dtScheduler.Rows.Count > 0)
                {
                    dt = dtScheduler.Copy();
                }
                int j = 0;
                for (int i = 0; i < dgScheduler.Rows.Count; i++)
                {
                    dt.Rows[j]["StartTime"] = dgScheduler.Rows[i].Cells[1].Value;
                    dt.Rows[j]["EndTime"] = dgScheduler.Rows[i].Cells[2].Value;
                    dt.Rows[j]["ExecutionTime"] = dgScheduler.Rows[i].Cells[3].Value;
                    dt.Rows[j]["Recursive"] = dgScheduler.Rows[i].Cells[4].Value;

                    clsDBOperation.UpdateRecords("insert into scheduler" + "(day,start_time,end_time,execution_time,`recursive`)" + " values('" + dt.Rows[j]["Day"].ToString() + "','" + dt.Rows[j]["StartTime"].ToString() + "','" + dt.Rows[j]["EndTime"].ToString() + "','" + dt.Rows[j]["ExecutionTime"].ToString() + "','" + dt.Rows[j]["Recursive"].ToString() + "')");
                    j++;
                    DisplayMessage("Record Save Successfully", "I");
                }
                ClsCommon.objSync.ReadSchedulerTime();
                ClsCommon.objSync.Refresh();

            Final:;

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmScheduler,Function :btnSave_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (ClsCommon.Mode == "Config")
            {
                Application.Exit();
            }
            else
            {
                ClsCommon.objScheduler.Hide();
                ClsCommon.objSync.ToolStripVisibility("Scheduler");
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
