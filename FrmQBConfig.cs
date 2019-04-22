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
using QbIntegration.QBClass;

namespace QbIntegration
{
    public partial class FrmQBConfig : Form
    {
        DataTable dtShippingItem = new DataTable();
        public FrmQBConfig()
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

        private void btnNext_Click(object sender, EventArgs e)
        {
            try
            {
                ClsCommon.objConfig.Hide();
                ClsCommon.objSync.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :btnNext_Click. Message:" + ex.Message);
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
                    ClsCommon.objConfig.Hide();
                    ClsCommon.objSync.ToolStripVisibility("QBConfig");
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :btnClose_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void FrmQBConfig_Load(object sender, EventArgs e)
        {
            try
            {
                LoadYear();

                LoadExistingData();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:frmQBConfig,Function :FrmQBConfig_Load. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void LoadYear()
        {
            try
            {
                DataTable dtYear = new DataTable();
                dtYear.Columns.Add("Year", typeof(string));
                dtYear.Columns.Add("Version", typeof(string));

                DataRow dr = dtYear.NewRow();
                dtYear.Rows.Add("--Select--", "0");
                dtYear.Rows.Add("2010", "9.0");
                dtYear.Rows.Add("2011", "10.0");
                dtYear.Rows.Add("2012", "10.0");
                dtYear.Rows.Add("2013", "12.0");
                dtYear.Rows.Add("2013+", "13.0");
                cmbYear.DataSource = dtYear;
                cmbYear.DisplayMember = "Year";
                cmbYear.ValueMember = "Version";
                cmbYear.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:frmQBConfig,Function :LoadYear. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        public Boolean StartQbSession()
        {
            Boolean isDone = false;
            try
            {
                if (Program.mySetting.QBFilePath != "")
                {
                    ClsCommon.QBCompanyFile = Program.mySetting.QBFilePath;
                }
                if (Program.mySetting.QBVersion != "")
                {
                    ClsCommon.QBVersion = Program.mySetting.QBVersion;
                }
                Program.dispose();
                CommonRef.oSession = null;
                if (CommonRef.oSession == null)
                {
                    ClsCommon.retMessage = QBConnection.OpenConnection_anyMode();
                    if (ClsCommon.retMessage["Status"].ToString() == "Session Created")
                        isDone = true;
                    else
                    {
                        ClsCommon.WriteErrorLogs("Start QBSession :" + ClsCommon.retMessage["Status"].ToString());
                        DisplayMessage(ClsCommon.retMessage["Status"].ToString(), "E");
                        DisplayMessage("Check QuickBook Session. Try to Connect", "I");
                        isDone = false;
                    }
                }
                return isDone;
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :StartQbSession. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
                return isDone;
            }
        }

        private void LoadData()
        {
            try
            {
                if (StartQbSession())
                {
                    //Load Account...
                    ClsCommon.retMessage = QBAccounts.Get_QBAccounts();
                    if (CommonRef.dtQBMaster.Rows.Count > 0)
                    {
                        //Income Acc
                        DataRow[] dr = CommonRef.dtQBMaster.Select("AccountType='Income'");
                        if (dr.Length > 0)
                        {
                            DataTable dtIncome = dr.CopyToDataTable();
                            if (dtIncome.Rows.Count > 0)
                            {
                                DataRow drAdd = dtIncome.NewRow();
                                drAdd["ListID"] = "0";
                                drAdd["Name"] = "---Select---";
                                dtIncome.Rows.InsertAt(drAdd, 0);
                                cmbIncome.DataSource = dtIncome;
                                cmbIncome.DisplayMember = "Name";
                                cmbIncome.ValueMember = "ListID";
                            }
                            else
                            {
                                cmbIncome.Items.Insert(0, "--Select--");
                            }
                            cmbIncome.SelectedIndex = 0;

                            //Cogs Acc
                            DataRow[] dr1 = CommonRef.dtQBMaster.Select("AccountType='CostOfGoodsSold'");
                            if (dr1.Length > 0)
                            {
                                DataTable dtCogs = dr1.CopyToDataTable();
                                if (dtCogs.Rows.Count > 0)
                                {
                                    DataRow drAdd = dtCogs.NewRow();
                                    drAdd["ListID"] = "0";
                                    drAdd["Name"] = "---Select---";
                                    dtCogs.Rows.InsertAt(drAdd, 0);
                                    cmbCogs.DataSource = dtCogs;
                                    cmbCogs.DisplayMember = "Name";
                                    cmbCogs.ValueMember = "ListID";
                                }
                                else
                                {
                                    cmbCogs.Items.Insert(0, "--Select--");
                                }
                                cmbCogs.SelectedIndex = 0;

                            }

                            //Asset Acc
                            DataRow[] dr2 = CommonRef.dtQBMaster.Select("[AccountType] like '%Asset%'");
                            if (dr2.Length > 0)
                            {
                                DataTable dtAsset = dr2.CopyToDataTable();
                                if (dtAsset.Rows.Count > 0)
                                {
                                    DataRow drAdd = dtAsset.NewRow();
                                    drAdd["ListID"] = "0";
                                    drAdd["Name"] = "---Select---";
                                    dtAsset.Rows.InsertAt(drAdd, 0);
                                    cmbAsset.DataSource = dtAsset;
                                    cmbAsset.DisplayMember = "Name";
                                    cmbAsset.ValueMember = "ListID";
                                }
                                else
                                {
                                    cmbAsset.Items.Insert(0, "--Select--");
                                }
                                cmbAsset.SelectedIndex = 0;

                            }
                        }

                    }

                    //TaxItem...
                    //cmbTaxItem.Items.Clear();
                    ////Tax Item
                    //ClsCommon.retMessage = QBSalesTaxItem.Retrieve_QBItemSalesTax();
                    //if (ClsCommon.retMessage["Status"].Contains("Error :") == false)
                    //{
                    //    if (ClsCommon.retMessage["Name"] != "" && ClsCommon.retMessage["Name"] != null)
                    //    {
                    //        string[] Data = ClsCommon.retMessage["Name"].Split(',');
                    //        foreach (string str in Data)
                    //        {
                    //            cmbTaxItem.Items.Add(str);
                    //        }
                    //    }
                    //}

                    Program.dispose();
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :LoadExistingData. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void LoadExistingData()
        {
            try
            {
                if (Program.mySetting.QBFilePath != "")
                {
                    LoadData();

                    txtQBPath.Text = Program.mySetting.QBFilePath;
                    txtVersion.Text = Program.mySetting.QBVersion;
                    cmbYear.Text = Program.mySetting.QBYear;

                    if (Program.mySetting.IncomeAcc != "")
                        cmbIncome.SelectedValue = Program.mySetting.IncomeAcc;

                    if (Program.mySetting.CogsAcc != "")
                        cmbCogs.SelectedValue = Program.mySetting.CogsAcc;

                    if (Program.mySetting.AssetAcc != "")
                        cmbAsset.SelectedValue = Program.mySetting.AssetAcc;

                    if (Program.mySetting.InvStartDate != "")
                    {
                        dtInvDate.Value = Convert.ToDateTime(Program.mySetting.InvStartDate);
                        chkInvDate.Checked = true;
                        
                    }

                    GrpItemSetting.Enabled = true;
                    GrpInvSetting.Enabled = true;
                    btnNext.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :LoadExistingData. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private Boolean CheckValidation()
        {
            Boolean ISValid = true;
            try
            {
                if (cmbYear.SelectedIndex == 0)
                {
                    ISValid = false;
                    DisplayMessage("Please Select Year", "E");
                    cmbYear.Focus();
                    goto Final;
                }

            Final:
                return ISValid;
            }
            catch (Exception ex)
            {
                return ISValid;
                ClsCommon.WriteErrorLogs("Function :CheckValidation. Message:" + ex.Message);
            }
        }

        private void btnQBAccess_Click(object sender, EventArgs e)
        {
            try
            {
                if (CheckValidation())
                {
                    lblErrorMsg.Text = "";
                    ClsCommon.retMessage.Clear();
                    Program.dispose();
                    ClsCommon.QBCompanyFile = "";
                    string OldQBPath = txtQBPath.Text;
                    ClsCommon.retMessage = QBConnection.OpenConnection_anyMode();
                    if (ClsCommon.retMessage["Status"].Contains("Error:") == false)
                    {
                        ClsCommon.retMessage = QBConnection.GetCompanyInfo(txtVersion.Text);
                        txtQBPath.Text = ClsCommon.retMessage["Path"].ToString();
                        txtVersion.Text = ClsCommon.retMessage["Version"].ToString();

                        DisplayMessage("QuickBooks Company File Permission set Successfully", "I");
                        ClsCommon.WriteErrorLogs("Function :ConfigureGetPath. Message:QuickBooks Company File Permission set Successfully");

                        Program.mySetting.QBFilePath = txtQBPath.Text;
                        Program.mySetting.QBVersion = txtVersion.Text;
                        Program.mySetting.QBYear = cmbYear.Text.ToString();
                        Program.mySetting.Save();

                        GrpItemSetting.Enabled = true;
                        GrpInvSetting.Enabled = true;


                        LoadData();
                        LoadExistingData();
                    }
                    else
                    {
                        DisplayMessage(ClsCommon.retMessage["Status"].ToString(), "E");
                        ClsCommon.WriteErrorLogs("Function :ConfigureGetPathChk. Message:" + ClsCommon.retMessage["Status"].ToString());
                    }
                    Program.dispose();
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :btnQBAccess_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }
        private Boolean CheckValidationForItem()
        {
            Boolean ISValid = true;
            try
            {
                if (cmbIncome.SelectedIndex == 0)
                {
                    ISValid = false;
                    DisplayMessage("Please Select IncomeAcc", "E");
                    cmbIncome.Focus();
                    goto Final;
                }
                else if (cmbCogs.SelectedIndex == 0)
                {
                    ISValid = false;
                    DisplayMessage("Please Select Cogs Acc", "E");
                    cmbCogs.Focus();
                    goto Final;
                }
                else if (cmbAsset.SelectedIndex == 0)
                {
                    ISValid = false;
                    DisplayMessage("Please Select Asset Acc", "E");
                    cmbAsset.Focus();
                    goto Final;
                }

            Final:
                return ISValid;
            }
            catch (Exception ex)
            {
                return ISValid;
                ClsCommon.WriteErrorLogs("Function :CheckValidation. Message:" + ex.Message);
            }
        }

        private void btnItemAccSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (CheckValidationForItem())
                {
                    Program.mySetting.IncomeAcc = cmbIncome.SelectedValue.ToString();
                    Program.mySetting.CogsAcc = cmbCogs.SelectedValue.ToString();
                    Program.mySetting.AssetAcc = cmbAsset.SelectedValue.ToString();
                    Program.mySetting.Save();
                    DisplayMessage("Item setting save successfully", "I");
                    btnNext.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :btnQBAccess_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void cmbYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbYear.SelectedIndex == 0)
                    txtVersion.Text = "";
                else
                    txtVersion.Text = cmbYear.SelectedValue.ToString();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :cmbYear_SelectedIndexChanged. Message:" + ex.Message);
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
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :Visibility. Message:" + ex.Message);
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

        private void btnInv_Click(object sender, EventArgs e)
        {
            try
            {
                if (chkInvDate.Checked == true)
                    Program.mySetting.InvStartDate = dtInvDate.Value.ToString("yyyy-MM-dd");
                else
                    Program.mySetting.InvStartDate = "";

               
                 
                    Program.mySetting.Save();

                    DisplayMessage("Invoice setting save successfully", "I");

                    btnNext.Enabled = true;
               
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBConfig,Function :btnInv_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void chkInvDate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkInvDate.Checked == true)
                dtInvDate.Enabled = true;
            else
                dtInvDate.Enabled = false;
        }
    }
}
