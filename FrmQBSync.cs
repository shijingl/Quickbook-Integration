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
using System.Threading;
using System.Timers;
using System.Collections.Specialized;
using System.Globalization;

namespace QbIntegration
{
    public partial class FrmQBSync : Form
    {
        System.Timers.Timer SyncTime1;
        //    System.Timers.Timer timer;
        Thread oThread;
        int Counter = 0;
        string LastSyncDate = "";
        DataTable dtchk = null;
        string Message = "";

        public FrmQBSync()
        {
            InitializeComponent();
            SyncTime1 = new System.Timers.Timer();
            SyncTime1.Elapsed += new ElapsedEventHandler(SyncTime1_Elapsed);
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

        private void btnSetScheduler_Click(object sender, EventArgs e)
        {
            try
            {
                ClsCommon.objScheduler.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :btnSetScheduler_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnViewLogs_Click(object sender, EventArgs e)
        {
            try
            {
                FrmAppsLog ObjApp = new FrmAppsLog();
                ObjApp.SetCurrentDate(DateTime.Now);
                ObjApp.ShowDialog();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :btnViewLogs_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            try
            {
                ClsCommon.objSync.Hide();
                ToolStripVisibility("QBSync");
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :btnHide_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncTime1_Elapsed(object sender, EventArgs e)
        {
            try
            {

                SyncTime1.Stop();
                SyncTime1.Enabled = false;

                for (int i = 0; i < dgSchedule.Rows.Count; i++)
                {
                    if (dgSchedule.Rows[i].Cells[3].Value.ToString() == "Waiting to Proceed as per Schedule")
                    {
                        DateTime dtSchedule = Convert.ToDateTime(dgSchedule.Rows[i].Cells[2].Value);
                        TimeSpan ts = dtSchedule - DateTime.Now;
                        if (ts.Hours == 0 && ts.Minutes == 0)
                        {
                            UpdateStatusMessage("Process Start", i);
                            if (dgSchedule.Rows[i].Cells[1].Value.ToString() == "Sync")
                            {

                                StartProcess();
                            }
                            UpdateStatusMessage("Process Complete", i);

                        }
                    }
                    Counter++;
                }
                if (Counter == 20)
                {
                    ReadSchedulerTime();
                    Counter = 0;
                }
                else
                {
                    SyncTime1.Enabled = true;
                    SyncTime1.Start();

                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncTime1_Elapsed. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void UpdateStatusMessage(string Msg, int i)
        {
            try
            {
                dgSchedule.Rows[i].Cells[3].Value = Msg;
                if (Msg.Contains("Start"))
                    dgSchedule.Rows[i].Cells[3].Style.ForeColor = Color.Orange;
                else
                    dgSchedule.Rows[i].Cells[3].Style.ForeColor = Color.Green;
                dgSchedule.Rows[i].Cells[3].Style.Font = new Font("Verdana", 10, FontStyle.Bold);
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :UpdateStatusMessage. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        public void ReadSchedulerTime()
        {
            try
            {
                dgSchedule.Rows.Clear();
                DisplayMessage("", "I");
                int Serial = 0;
                DataTable dtSchedule = new DataTable();
                dtSchedule = clsDBOperation.RetrieveDataforSync("SELECT * FROM scheduler  WHERE day='" + CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek).ToUpper() + "'");

                // int Counter = 0;
                foreach (DataRow dr in dtSchedule.Rows)
                {
                    string st = dr["start_time"].ToString();
                    if (dr["start_time"].ToString() == "" || dr["end_time"].ToString() == "" || dr["execution_time"].ToString() == "" || dr["start_time"].ToString() == "00:00:00" || dr["end_time"].ToString() == "00:00:00" || dr["execution_time"].ToString() == "00:00:00")
                    { }
                    else
                    {
                        DateTime dtStart = Convert.ToDateTime(dr["start_time"].ToString());
                        DateTime dtEnd = Convert.ToDateTime(dr["end_time"].ToString());
                        DateTime dtExecution = Convert.ToDateTime(dr["execution_time"].ToString());
                        Int32 Recursive = Convert.ToInt32(dr["recursive"].ToString());

                        while (dtExecution >= dtStart && dtExecution <= dtEnd)
                        {
                            dgSchedule.Rows.Add();
                            dgSchedule.Rows[Serial].Cells[0].Value = (Serial + 1);
                            dgSchedule.Rows[Serial].Cells[1].Value = "Sync";
                            dgSchedule.Rows[Serial].Cells[2].Value = dtExecution;
                            if (dtExecution < DateTime.Now)
                            {
                                dgSchedule.Rows[Serial].Cells[3].Value = "Process Complete";
                                dgSchedule.Rows[Serial].Cells[3].Style.ForeColor = Color.Green;
                                dgSchedule.Rows[Serial].Cells[3].Style.Font = new Font("Verdana", 10, FontStyle.Regular);
                            }
                            else
                            {
                                dgSchedule.Rows[Serial].Cells[3].Value = "Waiting to Proceed as per Schedule";
                                dgSchedule.Rows[Serial].Cells[3].Style.ForeColor = Color.Blue;
                                dgSchedule.Rows[Serial].Cells[3].Style.Font = new Font("Verdana", 10, FontStyle.Regular);
                            }
                            Serial += 1;
                            if (Recursive > 0)
                                dtExecution = dtExecution.AddMinutes(Recursive);
                            else
                                break;
                        }
                        break;
                    }
                }
                dgSchedule.Sort(dgSchedule.Columns[2], ListSortDirection.Ascending);
                SyncTime1.Enabled = true;
                SyncTime1.Interval = 50000;
                SyncTime1.Start();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :ReadSchedulerTime. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void FrmQBSync_Load(object sender, EventArgs e)
        {

            oThread = new Thread(new ThreadStart(StartProcess1));
            CheckForIllegalCrossThreadCalls = false;
            oThread.Start();
        }

        private void StartProcess1()
        {
            try
            {
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon1.BalloonTipText = "I am a NotifyIcon Balloon";
                notifyIcon1.ShowBalloonTip(1000);

                //if (Program.mySetting.LastSyncDate != "")
                //{
                //    lblLastSyncDate.Text = Program.mySetting.LastSyncDate;
                //}
                ReadSchedulerTime();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :StartProcess. Message:" + ex.Message);
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

        private void Clear()
        {
            lblNew.Text = "0";
            lblProcess.Text = "0";
            lblSkip.Text = "0";
            lblTotal.Text = "0";
            lblFail.Text = "0";
            lblUpdate.Text = "0";
        }

        private void StartProcess()
        {
            try
            {
                ClsCommon.WriteErrorLogs("*****************************Start*****************************");
                LastSyncDate = DateTime.Now.ToString();


                if (StartQbSession())
                {

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start Customers DBToQB*****************************");
                    DisplayMessage("Application try to Sync Customers From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync Customers From DB To QB");
                    SyncCustomerFromDBToQB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync Customers From DB To QB");
                    DisplayMessage("Application Complete to Sync Customers From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End Customers DBToQB*****************************");
                    WriteHistory("CustomerFromDBToQB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start Customers QBToDB*****************************");
                    DisplayMessage("Application try to Sync Customers From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync Customers From QB To DB");
                    SyncCustomerFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync Customers From QB To DB");
                    DisplayMessage("Application Complete to Sync Customers From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End Customers QBToDB*****************************");
                    WriteHistory("CustomerFromQBToDB");





                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start Items DBToQB*****************************");
                    DisplayMessage("Application try to Sync Items From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync Items From DB To QB");
                    SyncItemsFromDBToQB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync Items From DB To QB");
                    DisplayMessage("Application Complete to Sync Items From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End Items DBToQB*****************************");
                    WriteHistory("ItemsFromDBToQB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start ItemsInventory QBToDB*****************************");
                    DisplayMessage("Application try to Sync ItemsInventory From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync ItemsInventory From QB To DB");
                    SyncItemInventoryFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync ItemsInventory From QB To DB");
                    DisplayMessage("Application Complete to Sync ItemsInventory From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End ItemsInventory QBToDB*****************************");
                    WriteHistory("ItemInventoryFromQBToDB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start ItemNonInventory QBToDB*****************************");
                    DisplayMessage("Application try to Sync ItemNonInventory From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync ItemNonInventory From QB To DB");
                    SyncItemNonInventoryFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync ItemNonInventory From QB To DB");
                    DisplayMessage("Application Complete to Sync ItemNonInventory From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End ItemNonInventory QBToDB*****************************");
                    WriteHistory("ItemNonInventoryFromQBToDB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start ItemService QBToDB*****************************");
                    DisplayMessage("Application try to Sync ItemService From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync ItemService From QB To DB");
                    SyncItemServiceFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync ItemService From QB To DB");
                    DisplayMessage("Application Complete to Sync ItemService From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End ItemService QBToDB*****************************");
                    WriteHistory("ItemServiceFromQBToDB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start ItemOtherCharge QBToDB*****************************");
                    DisplayMessage("Application try to Sync ItemOtherCharge From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync ItemOtherCharge From QB To DB");
                    SyncItemOtherChargeFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync ItemOtherCharge From QB To DB");
                    DisplayMessage("Application Complete to Sync ItemOtherCharge From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End ItemOtherCharge QBToDB*****************************");
                    WriteHistory("ItemOtherChargeFromQBToDB");

                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start ItemSubTotal QBToDB*****************************");
                    DisplayMessage("Application try to Sync ItemSubTotal From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync ItemSubTotal From QB To DB");
                    SyncItemSubTotalFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync ItemSubTotal From QB To DB");
                    DisplayMessage("Application Complete to Sync ItemSubTotal From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End ItemSubTotal QBToDB*****************************");
                    WriteHistory("ItemSubTotalFromQBToDB");





                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start Invoice DBToQB*****************************");
                    DisplayMessage("Application try to Sync Invoice From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync Invoice From DB To QB");
                    SyncInvoiceFromDBToQB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync Invoice From DB To QB");
                    DisplayMessage("Application Complete to Sync Invoice From DB To QB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End Invoice DBToQB*****************************");
                    WriteHistory("InvoiceFromDBToQB");
                    Clear();
                    ClsCommon.WriteErrorLogs("*****************************Start Invoice QBToDB*****************************");
                    DisplayMessage("Application try to Sync Invoice From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Application try to Sync Invoice From QB To DB");
                    SyncInvoiceFromQBToDB();
                    ClsCommon.WriteErrorLogs("Application Complete to Sync Invoice From QB To DB");
                    DisplayMessage("Application Complete to Sync Invoice From QB To DB", "I");
                    ClsCommon.WriteErrorLogs("*****************************End Invoice QBToDB*****************************");
                    WriteHistory("InvoiceFromQBToDB");

                    Program.dispose();
                }

                Program.mySetting.LastSyncDate = LastSyncDate;
                Program.mySetting.Save();
                lblLastSyncDate.Text = Program.mySetting.LastSyncDate;

                ClsCommon.WriteErrorLogs("*****************************End*****************************");
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :StartProcess. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }


        #region DBToQB

        #region Customer
        private void SyncCustomerFromDBToQB()
        {
            try
            {
                lblMode.Text = "Customer DB To QB";

                DataTable dtCustomer = clsDBOperation.RetrieveDataforSync("SELECT id,first_name,last_name,email,is_active,fullname FROM auth_user where is_active=1 and ListID=''");
                if (dtCustomer.Rows.Count > 0)
                {
                    DisplayMessage("Sync All Customer from DB To QB", "I");

                    lblTotal.Text = (dtCustomer.Rows.Count).ToString();
                    foreach (DataRow dr in dtCustomer.Rows)
                    {
                        //First Check customer is exists or not in QB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();
                        string ListID = "";
                        string CustomerName = "";

                        if (dr["fullname"].ToString() != "")
                        {
                            CustomerName = dr["fullname"].ToString();
                            Boolean IsCustomer = CheckCustomerInQB(CustomerName, ref ListID);
                            if (IsCustomer)
                            {
                                NameValueCollection retData = new NameValueCollection();
                                retData.Add("Name", ClsCommon.RemoveSpecialCharacters(CustomerName, "Customer"));
                                retData.Add("FirstName", ClsCommon.RemoveSpecialCharacters(dr["first_name"].ToString(), "FirstName"));
                                retData.Add("LastName", ClsCommon.RemoveSpecialCharacters(dr["last_name"].ToString(), "LastName"));
                                retData.Add("Email", ClsCommon.RemoveSpecialCharacters(dr["email"].ToString(), ""));

                                retData.Add("IsActive", dr["is_active"].ToString().ToLower());

                                ClsCommon.retMessage = QBCustomer.AddCustomer(retData);
                                if (ClsCommon.retMessage["Status"].ToString() == "Customer Added")
                                {
                                    string Message = clsDBOperation.UpdateRecords("update auth_user set ListID=\"" + ClsCommon.retMessage["ListID"].ToString() + "\" where id=" + dr["id"].ToString() + "");

                                    lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Customer create successfully in QB; CustomerName:" + CustomerName);
                                    DisplayMessage("Customer create successfully in QB; CustomerName:" + CustomerName, "I");
                                }
                                else
                                {
                                    lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Error create Customer in QB; CustomerName:" + CustomerName + "; Message:" + ClsCommon.retMessage["Status"].ToString());
                                }
                            }
                            else
                            {
                                lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                                DisplayMessage("Customer is already exists in QuickBooks", "W");
                                ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncContactsFromZohoToQB;Message:Customer is already exists in QB;CustomerName:" + CustomerName);
                            }
                        }
                        else
                        {
                            lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                            ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncCustomerFromDBToQB. Message:Mandatory field is missing,FullName is blank");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncCustomerFromDBToQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private Boolean CheckCustomerInQB(string Name, ref string ListID)
        {
            Boolean IsCustomer = false;
            try
            {
                ClsCommon.retMessage = QBCustomer.CheckCustomer_Exists_ByName(ClsCommon.RemoveSpecialCharacters(Name.ToString(), "Customer"), "");
                if (ClsCommon.retMessage["Status"].ToString() == "Find Customer")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInQB. Message:Contact is find in QB");
                    IsCustomer = false;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No Customer Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInQB. Message:Contact is not found in QB,Now try to create in QB");
                    IsCustomer = true;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    IsCustomer = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return IsCustomer;
        }
        #endregion

        #region Item
        private void SyncItemsFromDBToQB()
        {
            try
            {
                lblMode.Text = "Items DB To QB";
                DataTable dtItem = clsDBOperation.RetrieveDataforSync("select * from order_item_table where status=1 and ListID=''");

                if (dtItem.Rows.Count > 0)
                {
                    DisplayMessage("Get All Items from DB To QB", "I");

                    lblTotal.Text = (dtItem.Rows.Count).ToString();
                    foreach (DataRow dr in dtItem.Rows)
                    {
                        //First Check Item is exists or not in QB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();
                        string ListID = ""; string ParentListID = "";
                        string ParentName = "";

                        Boolean otherItem = false; Boolean IsItem = false;
                        if (dr["name"].ToString() != "")
                        {
                            string ItemName = dr["name"].ToString();

                            IsItem = CheckItemInventory(ItemName, ref ListID, ParentName, ref ParentListID);

                            if (IsItem == false)
                            {

                                IsItem = CheckItemService(ItemName, ref ListID, ParentName, ref ParentListID);
                                otherItem = IsItem;
                            }

                            if (IsItem == false)
                            {
                                IsItem = CheckItemNonInventory(ItemName, ref ListID, ParentName, ref ParentListID);
                                otherItem = IsItem;
                            }
                            if (IsItem == false)
                            {
                                IsItem = CheckItemInventoryAssembly(ItemName, ref ListID, ParentName, ref ParentListID);
                                otherItem = IsItem;
                            }
                            if (IsItem == false)
                            {
                                IsItem = CheckItemOtherCharge(ItemName, ref ListID, ParentName, ref ParentListID);
                                otherItem = IsItem;
                            }
                            if (otherItem == true)
                            {
                                lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Item Name " + dr["Name"].ToString() + " in other type");

                            }
                            else
                            {
                                if (IsItem == false)
                                {
                                    IsItem = CreateItemInventory(dr, ref ListID, ParentListID);
                                    if (IsItem == true)
                                    {
                                        lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                        ClsCommon.WriteErrorLogs("Item create successfully; ItemName: " + ClsCommon.RemoveSpecialCharacters(dr["name"].ToString(), "Item"));
                                    }
                                    else
                                    {
                                        lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                        ClsCommon.WriteErrorLogs("Fail to create Item; ItemName: " + ClsCommon.RemoveSpecialCharacters(dr["name"].ToString(), "Item"));
                                    }
                                }
                                else
                                {
                                    string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + ListID + "\" where id=" + dr["id"].ToString() + "");

                                    lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Item is skip; ItemName: " + ClsCommon.RemoveSpecialCharacters(dr["name"].ToString().ToString(), "Item"));

                                }
                            }
                            if (IsItem == true && ListID == "")
                            {
                                ClsCommon.WriteErrorLogs("Erro: ItemName:" + ClsCommon.RemoveSpecialCharacters(dr["name"].ToString(), "Item"));
                                IsItem = false;
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();

                            }
                        }
                        else
                        {
                            ClsCommon.retMessage["Status"] = "ItemName Required!!!. ItemName is Blank";
                            lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                            IsItem = false;
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemsFromDBToQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private Boolean CheckItemInventory(string ItemName, ref string ListID, string ParentName, ref string ParentListID)
        {

            Boolean flag = false;
            try
            {
                if (ParentName != "")
                {
                    ClsCommon.retMessage = QBItemInventory.Check_QBItemInventory(ClsCommon.RemoveSpecialCharacters(ParentName, "Item"), "");
                    if (ClsCommon.retMessage["Status"].ToString() == "Find ItemInventory")
                    {
                        ParentListID = ClsCommon.retMessage["ListID"].ToString();
                        flag = true;
                    }
                    else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemInventory Found")
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryInQB. Message:Parent ItemInventory is not found in QB");
                    }
                }

                ClsCommon.retMessage = QBItemInventory.Check_QBItemInventory(ClsCommon.RemoveSpecialCharacters(ItemName.ToString(), "Item"), ParentListID);
                if (ClsCommon.retMessage["Status"].ToString() == "Find ItemInventory")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryInQB. Message:ItemInventory is find in QB");
                    flag = true;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemInventory Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryInQB. Message:ItemInventory is not found in QB,Now try to create in QB");
                    flag = false;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return flag;
        }

        private Boolean CheckItemService(string ItemName, ref string ListID, string ParentName, ref string ParentListID)
        {


            Boolean flag = false;
            try
            {
                if (ParentName != "")
                {
                    ClsCommon.retMessage = QBItemService.Check_QBItemService(ClsCommon.RemoveSpecialCharacters(ParentName, "Item"), "");
                    if (ClsCommon.retMessage["Status"].ToString() == "Find ItemService")
                    {
                        ParentListID = ClsCommon.retMessage["ListID"].ToString();
                        flag = true;
                    }
                    else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemService Found")
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemServiceInQB. Message:Parent ItemService is not found in QB");
                    }
                }
                ClsCommon.retMessage = QBItemService.Check_QBItemService(ClsCommon.RemoveSpecialCharacters(ItemName.ToString(), "Item"), ParentListID);
                if (ClsCommon.retMessage["Status"].ToString() == "Find ItemService")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemServiceInQB. Message:ItemService is find in QB");
                    flag = true;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemService Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemServiceInQB. Message:ItemService is not found in QB,Now try to create in QB");
                    flag = false;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemServiceInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemServiceInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return flag;
        }

        private Boolean CheckItemNonInventory(string ItemName, ref string ListID, string ParentName, ref string ParentListID)
        {
            Boolean flag = false;
            try
            {
                if (ParentName != "")
                {
                    ClsCommon.retMessage = QBItemNonInventory.Check_QBItemNonInventory(ClsCommon.RemoveSpecialCharacters(ParentName, "Item"), "");
                    if (ClsCommon.retMessage["Status"].ToString() == "Find ItemNonInventory")
                    {
                        ParentListID = ClsCommon.retMessage["ListID"].ToString();
                        flag = true;
                    }
                    else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemNonInventory Found")
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemNonInventoryInQB. Message:Parent ItemNonInventory is not found in QB");
                    }
                }
                ClsCommon.retMessage = QBItemNonInventory.Check_QBItemNonInventory(ClsCommon.RemoveSpecialCharacters(ItemName.ToString(), "Item"), ParentListID);
                if (ClsCommon.retMessage["Status"].ToString() == "Find ItemNonInventory")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemNonInventoryInQB. Message:ItemNonInventory is find in QB");
                    flag = true;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemNonInventory Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemNonInventoryInQB. Message:ItemNonInventory is not found in QB,Now try to create in QB");
                    flag = false;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemNonInventoryInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemNonInventoryInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return flag;
        }

        private Boolean CheckItemInventoryAssembly(string ItemName, ref string ListID, string ParentName, ref string ParentListID)
        {

            Boolean flag = false;
            try
            {
                if (ParentName != "")
                {
                    ClsCommon.retMessage = QBItemInventoryAssembly.Check_QBItemNonInventory(ClsCommon.RemoveSpecialCharacters(ParentName, "Item"), "");
                    if (ClsCommon.retMessage["Status"].ToString() == "Find ItemInventoryAssembly")
                    {
                        ParentListID = ClsCommon.retMessage["ListID"].ToString();
                        flag = true;
                    }
                    else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemInventoryAssembly Found")
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryAssemblyInQB. Message:Parent ItemInventoryAssembly is not found in QB");
                    }
                }
                ClsCommon.retMessage = QBItemInventoryAssembly.Check_QBItemNonInventory(ClsCommon.RemoveSpecialCharacters(ItemName.ToString(), "Item"), ParentListID);
                if (ClsCommon.retMessage["Status"].ToString() == "Find ItemInventoryAssembly")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryAssemblyInQB. Message:ItemInventoryAssembly is find in QB");
                    flag = true;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemInventoryAssembly Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryAssemblyInQB. Message:ItemInventoryAssembly is not found in QB,Now try to create in QB");
                    flag = false;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryAssemblyInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInventoryAssemblyInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return flag;

        }

        private Boolean CheckItemOtherCharge(string ItemName, ref string ListID, string ParentName, ref string ParentListID)
        {

            Boolean flag = false;
            try
            {
                if (ParentName != "")
                {
                    ClsCommon.retMessage = QBItemOtherCharge.Check_QBItemOtherCharge(ClsCommon.RemoveSpecialCharacters(ParentName, "Item"), "");
                    if (ClsCommon.retMessage["Status"].ToString() == "Find ItemOtherCharge")
                    {
                        ParentListID = ClsCommon.retMessage["ListID"].ToString();
                        flag = true;
                    }
                    else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemOtherCharge Found")
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemOtherChargeInQB. Message:Parent ItemOtherCharge is not found in QB");
                    }
                }
                ClsCommon.retMessage = QBItemOtherCharge.Check_QBItemOtherCharge(ClsCommon.RemoveSpecialCharacters(ItemName.ToString(), "Item"), ParentListID);
                if (ClsCommon.retMessage["Status"].ToString() == "Find ItemOtherCharge")
                {
                    ListID = ClsCommon.retMessage["ListID"].ToString();
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemOtherChargeInQB. Message:ItemOtherCharge is find in QB");
                    flag = true;
                }
                else if (ClsCommon.retMessage["Status"].ToString() == "Error: No ItemOtherCharge Found")
                {

                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemOtherChargeInQB. Message:ItemOtherCharge is not found in QB,Now try to create in QB");
                    flag = false;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemOtherChargeInQB. Message:" + ClsCommon.retMessage["Status"].ToString());
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemOtherChargeInQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
            return flag;
        }

        private Boolean CreateItemInventory(DataRow dr, ref string ListId, string ParentListId)
        {
            Boolean flag = false;
            try
            {

                NameValueCollection retData = new NameValueCollection();
                retData.Add("Name", ClsCommon.RemoveSpecialCharacters(dr["name"].ToString(), "Item"));
                retData.Add("SalesDesc", ClsCommon.RemoveSpecialCharacters(dr["description"].ToString(), ""));
                retData.Add("QuantityOnHand", dr["qty"].ToString());

                retData.Add("IsActive", dr["status"].ToString().ToLower());

                if (Program.mySetting.IncomeAcc != "")
                    retData.Add("IncomeAccountRefID", Program.mySetting.IncomeAcc);
                if (Program.mySetting.CogsAcc != "")
                    retData.Add("COGSAccountRefID", Program.mySetting.CogsAcc);
                if (Program.mySetting.AssetAcc != "")
                    retData.Add("AssetAccountRefID", Program.mySetting.AssetAcc);


                ClsCommon.retMessage = QBItemInventory.Add_QBItemInventory(retData);
                if (ClsCommon.retMessage["Status"].ToString() == "Add ItemInventory")
                {
                    string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + ClsCommon.retMessage["ListID"].ToString() + "\" where id=" + dr["id"].ToString() + "");

                    ListId = ClsCommon.retMessage["ListID"].ToString();
                    flag = true;
                    ClsCommon.WriteErrorLogs("Create Item in QB Successfully ItemName:" + ClsCommon.RemoveSpecialCharacters(dr["name"].ToString(), "Item"));
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemsInQB. Message:" + ClsCommon.retMessage["Status"].ToString() + " in other type");
                    flag = false;

                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Error :CreateItem Message : " + ex.Message);
            }
            return flag;
        }
        #endregion

        #region Invoice
        private void SyncInvoiceFromDBToQB()
        {
            try
            {
                lblMode.Text = "Invoice DB To QB";
                DataTable dtInvoice = clsDBOperation.RetrieveDataforSync("select o.id,o.order_number,o.notes,o.total_amount,o.customer_id,o.tax_amount,CAST(o.created_at AS datetime) as Date,c.first_name,c.fullname,c.last_name from order_order_table as o left join auth_user as c on c.ID=o.customer_id where o.txn_id=''");
                if (dtInvoice.Rows.Count > 0)
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync, Function:SyncInvoiceFromDBToQB,Total Invoice Count:" + dtInvoice.Rows.Count);

                    lblTotal.Text = (dtInvoice.Rows.Count).ToString();
                    ClsCommon.WriteErrorLogs("Total Invoice:" + dtInvoice.Rows.Count);
                    NameValueCollection retInvoice = new NameValueCollection();

                    foreach (DataRow dr in dtInvoice.Rows)
                    {
                        DisplayMessage("", "I");
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();
                        try
                        {
                            retInvoice.Clear();

                            //if (dtChk.Rows.Count == 0)
                            //{
                            ClsCommon.WriteErrorLogs("Application check Customers is exists in QB");
                            string CustomerListID = "";
                            string CustomerName = dr["fullname"].ToString();

                            if (dr["fullname"].ToString() != "")
                            {
                                Boolean isContacts = CheckCustomerInQB(CustomerName, ref CustomerListID);
                                if (isContacts == true && CustomerListID == "")
                                {
                                    lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Invoice Fail DB To QB Message : Customer Not Found ,Customer Name: " + CustomerName + "; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));
                                    goto Next;
                                }
                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Invoice Fail DB To QB Message :FirstName and LastName is blank of invoice; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));
                                goto Next;

                            }


                            ClsCommon.WriteErrorLogs("Application try to check Items exists in QB");
                            StringBuilder ItemInfo = new StringBuilder(); DataTable dtItem = new DataTable();
                            Boolean isItems = CheckItemsInQB(dr, ref ItemInfo, ref dtItem);
                            if (isItems == false)
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Invoice Fail Zoho To QB Message : Item Not Found ; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));

                                goto Next;
                            }


                            ClsCommon.WriteErrorLogs("Application try to add Invoice in QuickBook. InvoiceNo:" + dr["order_number"].ToString());
                            retInvoice.Add("CustomerRef", CustomerListID);
                            retInvoice.Add("ItemsInfo", ItemInfo.ToString());
                            retInvoice.Add("RefNumber", ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));

                            if (dr["Date"].ToString() != "")
                                retInvoice.Add("TxnDate", ClsCommon.ConvertDate_QBFormat(dr["Date"].ToString()));

                            retInvoice.Add("Memo", ClsCommon.RemoveSpecialCharacters(dr["notes"].ToString(), ""));

                            //if (Convert.ToDecimal(dr["Tax"].ToString()) > 0)
                            //    retInvoice.Add("ItemSalesTaxRef", Program.mysetting.TaxItem);
                            //else
                            //    retInvoice.Add("ItemSalesTaxRef", Program.mysetting.NonTaxItem);

                            //if (SalesTaxcode != "")
                            //    retInvoice.Add("CustomerSalesTaxCodeRef", SalesTaxcode);


                            ClsCommon.retMessage = QBInvoice.Add_QBInvoice(retInvoice);

                            if (ClsCommon.retMessage["Status"].ToString() == "Find Invoice")
                            {
                                string TxnID = CommonRef.dtQBMaster.Rows[0]["TxnID"].ToString();
                                string Message = clsDBOperation.UpdateRecords("update order_order_table set txn_id=\"" + TxnID + "\" where id=" + dr["id"].ToString() + "");

                                //Update TxnLineID in QB...

                                foreach (DataRow drItem in CommonRef.dtQBChild.Rows)
                                {
                                    DataRow[] dr1 = dtItem.Select("name='" + drItem["ItemFullName"].ToString() + "'");
                                    if (dr1.Length > 0)
                                    {
                                        DataTable dtID = dr1.CopyToDataTable();
                                        string Message1 = clsDBOperation.UpdateRecords("update order_orderitems_table set txn_id=\"" + drItem["TxnLineID"].ToString() + "\" where id=" + dtID.Rows[0]["id"].ToString() + "");
                                    }
                                }
                                lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Application Complete to add Invoice in QuickBook, InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));


                                //Create ReceivePayment in QB...
                                SyncReceivePaymentFromDBToQB(dr, ref CustomerListID, ref TxnID);

                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Fail to create Invoice in QB; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER") + ";Message:" + ClsCommon.retMessage["Status"].ToString());
                                DisplayMessage("Error:" + ClsCommon.retMessage["Status"].ToString(), "E");
                            }
                        //}
                        //else
                        //{
                        //    lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                        //    ClsCommon.WriteErrorLogs("Invoice Skip Zoho To QB. Message:Invoice is already exists in QB,InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER") + " InvoiceID:" + dr["id"].ToString());
                        //}

                        Next:
                            this.Refresh();
                        }
                        catch (Exception ex)
                        {
                            ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncInvoiceFromDBToQB. Message:" + ex.Message);

                        }
                    }

                }
                else
                {
                    Clear();
                    ClsCommon.WriteErrorLogs("Invoice row count is:0");
                }


            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncInvoiceFromDBToQB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private Boolean CheckItemsInQB(DataRow dr, ref StringBuilder ItemInfos, ref DataTable dtItem)
        {
            Boolean IsItem = false;
            try
            {

                DataTable dtInvoiceDetail = clsDBOperation.RetrieveDataforSync("select o.*,i.name from order_orderitems_table as o left join order_item_table as i on i.ID=o.item_id_id where order_id_id=" + dr["ID"].ToString() + "");
                for (int i = 0; i < dtInvoiceDetail.Rows.Count; i++)
                {
                    try
                    {

                        string Sales = "NON";
                        // Check Item is Exist in QB...
                        #region ItemInventory
                        string ListID = "";
                        string ParentName = "", ParentListID = "";
                        string ProductName = dtInvoiceDetail.Rows[i]["name"].ToString();

                        if (ProductName != "")
                        {
                            Boolean CheckFlag = CheckItemInventory(ProductName, ref ListID, ParentName, ref ParentListID);
                            if (CheckFlag == false)
                                CheckFlag = CheckItemService(dtInvoiceDetail.Rows[i]["name"].ToString(), ref ListID, ParentName, ref ParentListID);
                            if (CheckFlag == false)
                                CheckFlag = CheckItemNonInventory(dtInvoiceDetail.Rows[i]["name"].ToString(), ref ListID, ParentName, ref ParentListID);
                            if (CheckFlag == false)
                                CheckFlag = CheckItemOtherCharge(dtInvoiceDetail.Rows[i]["name"].ToString(), ref ListID, ParentName, ref ParentListID);
                            if (CheckFlag == false)
                                CheckFlag = CreateItemInventory(dtInvoiceDetail.Rows[i], ref ListID, ParentListID);

                            if (CheckFlag == true)
                            {
                                ClsCommon.WriteErrorLogs("Find Item in QB");
                                decimal Amount = 0;
                                if (dtInvoiceDetail.Rows[i]["amount"].ToString() != "")
                                    Amount = Convert.ToDecimal(dtInvoiceDetail.Rows[i]["amount"].ToString());

                                if (ItemInfos.ToString() == "")
                                    ItemInfos.Append(ListID + "=" + "" + "=" + dtInvoiceDetail.Rows[i]["qty"].ToString() + "=" + "" + "=" + Amount.ToString("N2").Replace(",", "") + "=====" + Sales);
                                else
                                    ItemInfos.Append("^" + ListID + "=" + "" + "=" + dtInvoiceDetail.Rows[i]["qty"].ToString() + "=" + "" + "=" + Amount.ToString("N2").Replace(",", "") + "=====" + Sales);

                                IsItem = true;
                            }
                            else
                            {
                                IsItem = false;
                                break;
                            }
                        }
                        else
                        {
                            IsItem = false;
                            break;
                        }

                        #endregion
                    }
                    catch (Exception ex)
                    {
                        ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemsInQB. Message:" + ex.Message);
                        IsItem = false;
                    }

                }
                dtItem = dtInvoiceDetail.Copy();

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemsInQB. Message:" + ex.Message);
                IsItem = false;
            }
            return IsItem;
        }

        private void SyncReceivePaymentFromDBToQB(DataRow dr, ref string CustomerListID, ref string TxnID)
        {

            try
            {
                DisplayMessage("Application try to create ReceivePayment in QB", "I");
                ClsCommon.WriteErrorLogs("Application try to create ReceivePayment in QB");

                DataTable dtPayment = clsDBOperation.RetrieveDataforSync("select id,order_id_id,order_item_id_id,item_id_id,qty,CAST(updated_at AS datetime) as Date from order_receiveitems_table where order_id_id=" + dr["id"].ToString() + "");
                if (dtPayment.Rows.Count > 0)
                {
                    NameValueCollection retPayment = new NameValueCollection();

                    retPayment.Clear();
                    if (dr["Date"].ToString() != "")
                        retPayment.Add("TxnDate", ClsCommon.ConvertDate_QBFormat(dtPayment.Rows[0]["Date"].ToString()));
                    retPayment.Add("CustomerListIDRef", CustomerListID);
                    retPayment.Add("RefNumber", ClsCommon.RemoveSpecialCharacters(dtPayment.Rows[0]["id"].ToString(), "REFNUMBER"));

                    retPayment.Add("AppliedToTxnAdd", "");
                    retPayment.Add("TxnID", TxnID);
                    decimal sum = 0;
                    foreach (DataRow drcount in CommonRef.dtQBChild.Rows)
                    {
                        sum += Convert.ToDecimal(drcount["Amount"]);
                    }
                    retPayment.Add("TotalAmount", sum.ToString());

                    ClsCommon.retMessage = QBReceivePayments.Add_QBReceiptPayments(retPayment);
                    if (ClsCommon.retMessage["Status"].ToString() == "Add ReceiptPayments")
                    {
                        string Message1 = clsDBOperation.UpdateRecords("update order_receiveitems_table set txn_id=\"" + ClsCommon.retMessage["TxnID"].ToString() + "\" where id=" + dtPayment.Rows[0]["id"].ToString() + "");

                        ClsCommon.WriteErrorLogs("Application Complete to add ReceivePayment of Invoice in QuickBook, InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER"));
                    }
                    else
                    {
                        ClsCommon.WriteErrorLogs("Fail to create ReceivePayment of Invoice in QB; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER") + ";Message:" + ClsCommon.retMessage["Status"].ToString());
                        DisplayMessage("Fail to create ReceivePayment of Invoice in QB; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER") + ";Message:" + ClsCommon.retMessage["Status"].ToString(), "E");
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncReceivePaymentFromDBToQB; InvoiceNo:" + ClsCommon.RemoveSpecialCharacters(dr["order_number"].ToString(), "REFNUMBER") + "; Message:" + ex.Message);
            }
        }

        #endregion

        #endregion





        #region QBToDB
        private void SyncCustomerFromQBToDB()
        {
            try
            {
                lblMode.Text = "Customer QB To DB";
                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBCustomer.Retrieve_Customers_Full("", "");
                else
                    ClsCommon.retMessage = QBCustomer.Retrieve_Customers_Full(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));

                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All Customer from QB To DB", "I");
                    int IsActive = 0;
                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check customer is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();

                        IsActive = 0;
                        if (dr["IsActive"].ToString() == "true")
                            IsActive = 1;
                        string UserName = String.Format("{0:d9}", (DateTime.Now.Ticks / 10) % 1000000000);

                        dtchk = clsDBOperation.RetrieveDataforSync("select ID,first_name,last_name,fullname from auth_user where ListID=\"" + dr["ListID"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            dtchk = clsDBOperation.RetrieveDataforSync("select ID,first_name,last_name,fullname from auth_user where fullname=\"" + dr["Name"].ToString() + "\"");
                            if (dtchk.Rows.Count == 0)
                            {
                                Message = clsDBOperation.UpdateRecords("insert into auth_user" + "(username,first_name,last_name,email,is_active,ListID,fullname)" + " values(\"" + UserName + "\",\"" + dr["FirstName"].ToString().Replace("&amp;", "&") + "\",\"" + dr["LastName"].ToString().Replace("&amp;", "&") + "\",\"" + dr["Email"].ToString().Replace("&amp;", "&") + "\"," + IsActive + ",\"" + dr["ListID"].ToString() + "\",\"" + dr["Name"].ToString().Replace("&amp;", "&") + "\")");

                                if (Message == "Operation successfully")
                                {
                                    lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Customer create successfully in DB; FullName:" + dr["Name"].ToString());
                                    DisplayMessage("Customer create successfully in DB;FullName:" + dr["Name"].ToString(), "I");
                                }
                                else
                                {
                                    lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Error create Customer in DB; FullName:" + dr["Name"].ToString() + "; Message:" + Message);
                                }

                            }
                            else
                            {
                                string Message = clsDBOperation.UpdateRecords("update auth_user set ListID=\"" + dr["ListID"].ToString() + "\",fullname=\"" + dr["Name"].ToString() + "\",first_name=\"" + dr["FirstName"].ToString() + "\",last_name=\"" + dr["LastName"].ToString() + "\",is_active=\"" + IsActive + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");

                                lblUpdate.Text = (Convert.ToInt16(lblUpdate.Text) + 1).ToString();
                                DisplayMessage("Customer is already exists in DB so it Update", "W");
                                ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncCustomerFromQBToDB;Message:Customer is already exists in DB so it Update;FullName:" + dr["Name"].ToString());
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update auth_user set ListID=\"" + dr["ListID"].ToString() + "\",fullname=\"" + dr["Name"].ToString() + "\",first_name=\"" + dr["FirstName"].ToString() + "\",last_name=\"" + dr["LastName"].ToString() + "\",is_active=\"" + IsActive + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");

                            lblUpdate.Text = (Convert.ToInt16(lblUpdate.Text) + 1).ToString();
                            DisplayMessage("Customer is already exists in DB so it Update", "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncCustomerFromQBToDB;Message:Customer is already exists in DB so it Update;FullName:" + dr["Name"].ToString());
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncCustomerFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncItemInventoryFromQBToDB()
        {
            try
            {
                lblMode.Text = "ItemInventory QB To DB";

                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBItemInventory.Retrieve_QBItemInventory("", "");
                else
                    ClsCommon.retMessage = QBItemInventory.Retrieve_QBItemInventory(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));

                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All ItemInventory from QB To DB", "I");
                    int IsActive = 0;
                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    ClsCommon.WriteErrorLogs("Total ItemInventory count is:" + CommonRef.dtQBMaster.Rows.Count);
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check Item is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();
                        IsActive = 0;
                        if (dr["IsActive"].ToString() == "true")
                            IsActive = 1;
                        dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where ListID=\"" + dr["ListID"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["name"].ToString() + "\"");
                            if (dtchk.Rows.Count == 0)
                            {


                                string Message = clsDBOperation.UpdateRecords("insert into order_item_table" + "(name,description,qty,status,ListID)" + " values(\"" + dr["name"].ToString() + "\",\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\",'" + dr["QuantityOnHand"].ToString() + "'," + IsActive + ",'" + dr["ListID"].ToString() + "')");

                                if (Message == "Operation successfully")
                                {
                                    lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("ItemInventory create successfully in DB; ItemName:" + dr["name"].ToString());
                                    DisplayMessage("ItemInventory create successfully in DB;ItemName:" + dr["name"].ToString(), "I");
                                }
                                else
                                {
                                    lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                    ClsCommon.WriteErrorLogs("Error create ItemInventory in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                                }
                            }
                            else
                            {
                                string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\",name=\"" + dr["Name"].ToString() + "\",status=\"" + IsActive + "\",description=\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");
                                lblUpdate.Text = (Convert.ToInt16(lblUpdate.Text) + 1).ToString();
                                DisplayMessage("ItemInventory is already exists in DB so it Update;ItemName:" + dr["name"].ToString(), "W");
                                ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemInventoryFromQBToDB;Message:ItemInventory is already exists in DB so it update;ItemName:" + dr["name"].ToString());
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\",name=\"" + dr["Name"].ToString() + "\",status=\"" + IsActive + "\",description=\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");
                            lblUpdate.Text = (Convert.ToInt16(lblUpdate.Text) + 1).ToString();
                            DisplayMessage("ItemInventory is already exists in DB so it Update;ItemName:" + dr["name"].ToString(), "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemInventoryFromQBToDB;Message:ItemInventory is already exists in DB so it update;ItemName:" + dr["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemInventoryFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncItemNonInventoryFromQBToDB()
        {
            try
            {
                lblMode.Text = "ItemNonInventory QB To DB";
                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBItemNonInventory.Retrieve_QBItemNonInventory("", "");
                else
                    ClsCommon.retMessage = QBItemNonInventory.Retrieve_QBItemNonInventory(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));

                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All ItemNonInventory from QB To DB", "I");

                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    ClsCommon.WriteErrorLogs("Total ItemNonInventory count is:" + CommonRef.dtQBMaster.Rows.Count);
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check Item is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();


                        DataTable dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["name"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            int IsActive = 0;
                            if (dr["IsActive"].ToString() == "true")
                                IsActive = 1;

                            string Message = clsDBOperation.UpdateRecords("insert into order_item_table" + "(name,description,qty,status,ListID)" + " values(\"" + dr["name"].ToString() + "\",\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\",'" + dr["QuantityOnHand"].ToString() + "'," + IsActive + ",'" + dr["ListID"].ToString() + "')");

                            if (Message == "Operation successfully")
                            {
                                lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("ItemNonInventory create successfully in DB; ItemName:" + dr["name"].ToString());
                                DisplayMessage("ItemNonInventory create successfully in DB;ItemName:" + dr["name"].ToString(), "I");
                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Error create ItemNonInventory in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");


                            lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                            DisplayMessage("ItemNonInventory is already exists in DB;ItemName:" + dr["name"].ToString(), "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemNonInventoryFromQBToDB;Message:ItemNonInventory is already exists in DB;ItemName:" + dr["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemNonInventoryFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncItemServiceFromQBToDB()
        {
            try
            {
                lblMode.Text = "ItemService QB To DB";


                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBItemService.Retrieve_QBItemService("", "");
                else
                    ClsCommon.retMessage = QBItemService.Retrieve_QBItemService(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));

                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All ItemService from QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Total ItemService count is:" + CommonRef.dtQBMaster.Rows.Count);
                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();

                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check Item is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();


                        DataTable dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["name"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            int IsActive = 0;
                            if (dr["IsActive"].ToString() == "true")
                                IsActive = 1;

                            string Message = clsDBOperation.UpdateRecords("insert into order_item_table" + "(name,description,qty,status,ListID)" + " values(\"" + dr["name"].ToString() + "\",\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\",'" + dr["QuantityOnHand"].ToString() + "'," + IsActive + ",'" + dr["ListID"].ToString() + "')");

                            if (Message == "Operation successfully")
                            {
                                lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("ItemService create successfully in DB; ItemName:" + dr["name"].ToString());
                                DisplayMessage("ItemService create successfully in DB;ItemName:" + dr["name"].ToString(), "I");
                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Error create ItemService in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");


                            lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                            DisplayMessage("ItemService is already exists in DB;ItemName:" + dr["name"].ToString(), "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemServiceFromQBToDB;Message:ItemService is already exists in DB;ItemName:" + dr["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemServiceFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncItemOtherChargeFromQBToDB()
        {
            try
            {
                lblMode.Text = "ItemOtherCharge QB To DB";

                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBItemOtherCharge.Retrieve_QBItemOtherCharge("", "");
                else
                    ClsCommon.retMessage = QBItemOtherCharge.Retrieve_QBItemOtherCharge(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));


                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All ItemOtherCharge from QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Total ItemOtherCharge count is:" + CommonRef.dtQBMaster.Rows.Count);
                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check Item is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();


                        DataTable dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["name"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            int IsActive = 0;
                            if (dr["IsActive"].ToString() == "true")
                                IsActive = 1;

                            string Message = clsDBOperation.UpdateRecords("insert into order_item_table" + "(name,description,qty,status,ListID)" + " values(\"" + dr["name"].ToString() + "\",\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\",'" + dr["QuantityOnHand"].ToString() + "'," + IsActive + ",'" + dr["ListID"].ToString() + "')");

                            if (Message == "Operation successfully")
                            {
                                lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("ItemOtherCharge create successfully in DB; ItemName:" + dr["name"].ToString());
                                DisplayMessage("ItemOtherCharge create successfully in DB;ItemName:" + dr["name"].ToString(), "I");
                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Error create ItemOtherCharge in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");


                            lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                            DisplayMessage("ItemOtherCharge is already exists in DB;ItemName:" + dr["name"].ToString(), "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemOtherChargeFromQBToDB;Message:ItemOtherCharge is already exists in DB;ItemName:" + dr["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemOtherChargeFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncItemSubTotalFromQBToDB()
        {
            try
            {
                lblMode.Text = "ItemSubTotal QB To DB";

                if (Program.mySetting.LastSyncDate == "")
                    ClsCommon.retMessage = QBItemSubTotal.Retrieve_QBItemSubTotal("", "");
                else
                    ClsCommon.retMessage = QBItemSubTotal.Retrieve_QBItemSubTotal(ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate), ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToShortDateString()));

                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {
                    DisplayMessage("Sync All ItemSubTotal from QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Total ItemSubTotal count is:" + CommonRef.dtQBMaster.Rows.Count);
                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check Item is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();


                        DataTable dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["name"].ToString() + "\"");
                        if (dtchk.Rows.Count == 0)
                        {
                            int IsActive = 0;
                            if (dr["IsActive"].ToString() == "true")
                                IsActive = 1;

                            string Message = clsDBOperation.UpdateRecords("insert into order_item_table" + "(name,description,qty,status,ListID)" + " values(\"" + dr["name"].ToString() + "\",\"" + dr["SalesDescription"].ToString().Replace("\"", "'") + "\",'" + dr["QuantityOnHand"].ToString() + "'," + IsActive + ",'" + dr["ListID"].ToString() + "')");

                            if (Message == "Operation successfully")
                            {
                                lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("ItemSubTotal create successfully in DB; ItemName:" + dr["name"].ToString());
                                DisplayMessage("ItemSubTotal create successfully in DB;ItemName:" + dr["name"].ToString(), "I");
                            }
                            else
                            {
                                lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Error create ItemSubTotal in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                            }
                        }
                        else
                        {
                            string Message = clsDBOperation.UpdateRecords("update order_item_table set ListID=\"" + dr["ListID"].ToString() + "\" where id=" + dtchk.Rows[0]["id"].ToString() + "");


                            lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                            DisplayMessage("ItemSubTotal is already exists in DB;ItemName:" + dr["name"].ToString(), "W");
                            ClsCommon.WriteErrorLogs("Form: FrmQBSync, Function: SyncItemSubTotalFromQBToDB;Message:ItemSubTotal is already exists in DB;ItemName:" + dr["name"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncItemSubTotalFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void SyncInvoiceFromQBToDB()
        {
            try
            {
                lblMode.Text = "Invoice QB To DB";
                NameValueCollection retInvoice = new NameValueCollection();
                if (Program.mySetting.LastSyncDate != "")
                {
                    retInvoice.Add("FromModifiedDate", ClsCommon.ConvertDate_QBFormat(Program.mySetting.LastSyncDate));
                    retInvoice.Add("ToModifiedDate", ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToString()));
                }
                else if (Program.mySetting.InvStartDate != "")
                {
                    retInvoice.Add("FromModifiedDate", ClsCommon.ConvertDate_QBFormat(Program.mySetting.InvStartDate));
                    retInvoice.Add("ToModifiedDate", ClsCommon.ConvertDate_QBFormat(DateTime.Now.ToString()));
                }

                ClsCommon.retMessage = QBInvoice.Retrieve_QBInvoice_Full(retInvoice);
                if (CommonRef.dtQBMaster.Rows.Count > 0)
                {

                    DisplayMessage("Sync All Invoice from QB To DB", "I");
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync, Function:SyncInvoiceFromQBToDB,Total Invoice Count:" + CommonRef.dtQBMaster.Rows.Count);

                    lblTotal.Text = (CommonRef.dtQBMaster.Rows.Count).ToString();
                    foreach (DataRow dr in CommonRef.dtQBMaster.Rows)
                    {
                        //First Check customer is exists or not in DB...If not then create it..
                        lblProcess.Text = (Convert.ToInt16(lblProcess.Text) + 1).ToString();
                        try
                        {
                            ClsCommon.WriteErrorLogs("Application try to create Invoice in DB");
                            string CustomerID = "";
                            DataTable dtChk = clsDBOperation.RetrieveDataforSync("select id,txn_id from order_order_table where txn_id='" + dr["TxnID"].ToString() + "'");
                            if (dtChk.Rows.Count == 0)
                            {
                                Boolean IsCustomer = CheckCustomerInDB(dr["CustomerFullName"].ToString(), ref CustomerID);
                                if (CustomerID != "")
                                {
                                    //Check Item in DB....
                                    StringBuilder ItemInfo = new StringBuilder(); DataTable dtItem = new DataTable();
                                    Boolean IsItem = CheckItemInDB(dr, ref ItemInfo);
                                    if (IsItem == false)
                                    {
                                        lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                        goto Next;
                                    }


                                    string Message = clsDBOperation.UpdateRecords("insert into order_order_table" + "(order_number,notes,created_at,total_amount,customer_id,txn_id,tax_amount)" + " values(\"" + dr["RefNumber"].ToString() + "\",\"" + dr["Memo"].ToString().Replace("\"", "'") + "\",'" + dr["TxnDate"].ToString() + "'," + dr["Subtotal"].ToString() + "," + CustomerID + ",'" + dr["TxnID"].ToString() + "'," + dr["SalesTaxTotal"].ToString() + ")");
                                    if (Message == "Operation successfully")
                                    {
                                        DataTable dtID = clsDBOperation.RetrieveDataforSync("select id from order_order_table order by id desc limit 1");
                                        if (dtID.Rows.Count > 0)
                                        {
                                            decimal Qty = 0;
                                            string[] Items = ItemInfo.ToString().Split('^');
                                            foreach (string Item in Items)
                                            {
                                                string[] Data = Item.Split('=');
                                                if (Data[1].ToString() != "")
                                                    Qty += Convert.ToDecimal(Data[1].ToString());

                                                string Message1 = clsDBOperation.UpdateRecords("insert into order_orderitems_table" + "(order_id_id,item_id_id,txn_id,qty,rate,amount)" + " values(" + dtID.Rows[0]["id"].ToString() + "," + Data[0].ToString() + ",'" + Data[4].ToString() + "'," + Data[1].ToString() + "," + Data[2].ToString() + "," + Data[3].ToString() + ")");
                                                if (Message1 == "Operation successfully")
                                                { }
                                                else
                                                {
                                                    ClsCommon.WriteErrorLogs("Error create InvoiceItems in DB; order_number:" + dr["RefNumber"].ToString() + "; Message:" + Message1);
                                                    goto Next;
                                                }
                                            }

                                            ClsCommon.WriteErrorLogs("Invoice create successfully in DB; order_number:" + dr["RefNumber"].ToString());
                                            DisplayMessage("Item create successfully in DB;order_number:" + dr["RefNumber"].ToString(), "I");
                                            lblNew.Text = (Convert.ToInt16(lblNew.Text) + 1).ToString();


                                            //Create ReceivePayment in DB...
                                            SyncReceivePaymentFromQBToDB(dr, ref dtID, ref Qty);


                                        }
                                    }
                                    else
                                    {
                                        lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                        ClsCommon.WriteErrorLogs("Error create Invoice in DB; ItemName:" + dr["name"].ToString() + "; Message:" + Message);
                                    }

                                }
                                else
                                {
                                    lblFail.Text = (Convert.ToInt16(lblFail.Text) + 1).ToString();
                                }
                            }
                            else
                            {
                                lblSkip.Text = (Convert.ToInt16(lblSkip.Text) + 1).ToString();
                                ClsCommon.WriteErrorLogs("Invoice Skip QB To DB,Message:Invoice is already exists in DB,OrderNumber:" + dr["RefNumber"].ToString());

                            }

                        }
                        catch (Exception ex)
                        {
                            ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncInvoiceFromQBToDB. Message:" + ex.Message);

                        }

                    Next:
                        this.Refresh();
                    }
                }

            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncInvoiceFromQBToDB. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private Boolean CheckCustomerInDB(string Name, ref string CustomerId)
        {
            bool flag = false;
            try
            {
                //string FirstName = ""; string LastName = "";

                //if (Name.ToString().Contains(" ") == true)
                //{
                //    var firstSpaceIndex = Name.ToString().IndexOf(" ");
                //    FirstName = Name.ToString().Substring(0, firstSpaceIndex);
                //    LastName = Name.ToString().Substring(firstSpaceIndex + 1);
                //}
                //else
                //{
                //    FirstName = Name.ToString();
                //}

                DataTable dtchk = clsDBOperation.RetrieveDataforSync("select id,first_name,last_name from auth_user where fullname=\"" + Name + "\"");
                if (dtchk.Rows.Count > 0)
                {
                    CustomerId = dtchk.Rows[0]["id"].ToString();
                    flag = true;
                }
                else
                {
                    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInDB. Message:Customer is not found in DB; fullname:" + Name);
                    flag = false;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckCustomerInDB. Message:" + ex.Message);
                flag = false;
            }
            return flag;
        }

        private Boolean CheckItemInDB(DataRow drInvoice, ref StringBuilder ItemInfo)
        {
            bool IsItem = false;
            try
            {
                DataRow[] dr1 = CommonRef.dtQBChild.Select("TxnID='" + drInvoice["TxnID"].ToString() + "'");
                if (dr1.Length > 0)
                {
                    DataTable dtDetail = dr1.CopyToDataTable();
                    foreach (DataRow dr in dtDetail.Rows)
                    {
                        if (dr["ItemFullName"].ToString() != "")
                        {
                            DataTable dtchk = clsDBOperation.RetrieveDataforSync("select * from order_item_table where name=\"" + dr["ItemFullName"].ToString() + "\"");
                            if (dtchk.Rows.Count > 0)
                            {
                                decimal Amount = 0; decimal Rate = 0;
                                if (dr["Amount"].ToString() != "")
                                    Amount = Convert.ToDecimal(dr["Amount"].ToString());

                                if (dr["Rate"].ToString() != "")
                                    Rate = Convert.ToDecimal(dr["Rate"].ToString());

                                if (ItemInfo.ToString() == "")
                                    ItemInfo.Append(dtchk.Rows[0]["id"].ToString() + "=" + dr["Quantity"].ToString() + "=" + Rate + "=" + Amount + "=" + dr["TxnLineID"].ToString());
                                else
                                    ItemInfo.Append("^" + dtchk.Rows[0]["id"].ToString() + "=" + dr["Quantity"].ToString() + "=" + Rate + "=" + Amount + "=" + dr["TxnLineID"].ToString());

                                IsItem = true;
                            }
                            else
                            {
                                if (Convert.ToDecimal(dr["Amount"].ToString()) > 0)
                                {
                                    ClsCommon.WriteErrorLogs("Erro:Function :CheckItemInDB. Message:Item is not found in DB; ItemName:" + dr["ItemFullName"].ToString());
                                    IsItem = false;
                                    goto Final;
                                }
                            }
                        }
                        else
                        {

                            if (dr["ItemFullName"].ToString() == "" && Convert.ToDecimal(dr["Amount"].ToString()) > 0)
                            {
                                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInDB. Message:Item is blank");
                                IsItem = false;
                                goto Final;
                            }
                            //else if(dr["ItemFullName"].ToString() == "" && Convert.ToDecimal(dr["Amount"].ToString()) == 0)
                            //{
                            //    ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInDB. Message:Item is blank");
                            //    IsItem = false;

                            //}
                        }
                    }
                Final:
                    this.Refresh();

                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :CheckItemInDB. Message:" + ex.Message);
                IsItem = false;
            }
            return IsItem;
        }

        private void SyncReceivePaymentFromQBToDB(DataRow dr, ref DataTable dtID, ref decimal Qty)
        {

            try
            {
                if (CommonRef.dtPayment.Rows.Count > 0)
                {


                    DataRow[] drPayment = CommonRef.dtPayment.Select("TxnID='" + dr["TxnID"].ToString() + "'");
                    if (drPayment.Length > 0)
                    {
                        DataTable dtPayment = drPayment.CopyToDataTable();
                        if (dtPayment.Rows.Count > 0)
                        {
                            DisplayMessage("Application try to create ReceivePayment in DB", "I");
                            ClsCommon.WriteErrorLogs("Application try to create ReceivePayment in DB");

                            //foreach (DataRow dr1 in dtPayment.Rows)
                            //{
                            string Message2 = clsDBOperation.UpdateRecords("insert into order_receiveitems_table" + "(order_id_id,txn_id,updated_at,qty)" + " values(" + dtID.Rows[0]["id"].ToString() + ",'" + dtPayment.Rows[0]["PaymentTxnID"].ToString() + "','" + dtPayment.Rows[0]["Date"].ToString() + "'," + Qty + ")");
                            if (Message2 == "Operation successfully")
                            {
                                ClsCommon.WriteErrorLogs("ReceivePayment create successfully in DB; Order_id:" + dtID.Rows[0]["id"].ToString());
                                DisplayMessage("ReceivePayment create successfully in DB; Order_id:" + dtID.Rows[0]["id"].ToString(), "I");
                            }
                            else
                            {
                                ClsCommon.WriteErrorLogs("Error create ReceivePayment in DB; Order_id:" + dtID.Rows[0]["id"].ToString() + "; Message:" + Message2);
                            }
                            //}
                        }
                    }
                    else
                    {
                        DisplayMessage("No ReceivePayment of this invoice from QB", "I");
                        ClsCommon.WriteErrorLogs("No ReceivePayment of this invoice from QB");
                    }
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :SyncReceivePaymentFromQBToDB; Message:" + ex.Message);
            }
        }


        #endregion

        private void mySqlDatabaseConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                mySqlDatabaseConfigurationToolStripMenuItem.Enabled = false;
                mySqlDatabaseConfigurationToolStripMenuItem.Checked = true;
                ClsCommon.objSqlConfig.Show();
                ClsCommon.objSqlConfig.Visibility();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :mySqlDatabaseConfigurationToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void quickBooksConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                quickBooksConfigurationToolStripMenuItem.Enabled = false;
                quickBooksConfigurationToolStripMenuItem.Checked = true;
                ClsCommon.objConfig.Show();
                ClsCommon.objConfig.Visibility();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :mySqlDatabaseConfigurationToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void showQBSyncToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                showQBSyncToolStripMenuItem.Enabled = false;
                showQBSyncToolStripMenuItem.Checked = true;

                hiddenQBSyncToolStripMenuItem.Enabled = true;
                hiddenQBSyncToolStripMenuItem.Checked = false;
                ClsCommon.objSync.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :showQBSyncToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void hiddenQBSyncToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                hiddenQBSyncToolStripMenuItem.Enabled = false;
                hiddenQBSyncToolStripMenuItem.Checked = true;

                showQBSyncToolStripMenuItem.Enabled = true;
                showQBSyncToolStripMenuItem.Checked = false;

                ClsCommon.objSync.Hide();


            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :hiddenQBSyncToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void schedulerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                schedulerToolStripMenuItem.Enabled = false;
                schedulerToolStripMenuItem.Checked = true;
                ClsCommon.objScheduler.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :schedulerToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }


        public void ToolStripVisibility(string Mode)
        {
            try
            {
                switch (Mode)
                {
                    case "SQLConfig":
                        mySqlDatabaseConfigurationToolStripMenuItem.Enabled = true;
                        mySqlDatabaseConfigurationToolStripMenuItem.Checked = false;
                        break;
                    case "QBConfig":
                        quickBooksConfigurationToolStripMenuItem.Enabled = true;
                        quickBooksConfigurationToolStripMenuItem.Checked = false;
                        break;
                    case "QBSync":
                        showQBSyncToolStripMenuItem.Enabled = true;
                        showQBSyncToolStripMenuItem.Checked = false;
                        break;
                    case "Scheduler":
                        schedulerToolStripMenuItem.Enabled = true;
                        schedulerToolStripMenuItem.Checked = false;
                        break;
                    case "History":
                        historyToolStripMenuItem.Enabled = true;
                        historyToolStripMenuItem.Checked = false;
                        break;
                }
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :ToolStripVisibility. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void WriteHistory(string type)
        {
            clsDBOperation.UpdateRecords("INSERT INTO history(TranType,TotalRec,InsertRec,UpdateRec,SkipRec,FailRec,CreatedDate)VALUES('" + type + "'," + Convert.ToInt32(lblTotal.Text) + "," + Convert.ToInt32(lblNew.Text) + "," + Convert.ToInt32(lblUpdate.Text) + "," + Convert.ToInt32(lblSkip.Text) + "," + Convert.ToInt32(lblFail.Text) + ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
        }

        private void btnHistory_Click(object sender, EventArgs e)
        {
            try
            {
                ClsCommon.ObjHistory.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :btnViewLogs_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void historyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                historyToolStripMenuItem.Enabled = false;
                historyToolStripMenuItem.Checked = true;
                ClsCommon.ObjHistory.Show();
            }
            catch (Exception ex)
            {
                ClsCommon.WriteErrorLogs("Form:FrmQBSync,Function :historyToolStripMenuItem_Click. Message:" + ex.Message);
                DisplayMessage("Error:" + ex.Message, "E");
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
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
