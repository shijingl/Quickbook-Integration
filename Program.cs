using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using QbIntegration.Properties;
using QbIntegration.clsHelper;
using QbIntegration.QBClass;

namespace QbIntegration
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static Settings mySetting = new Settings();

        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                CloseInstance();
                if (Program.mySetting.Connection == "")
                {
                    ClsCommon.Mode = "Config";
                    Application.Run(ClsCommon.objSqlConfig);
                }
                else if (Program.mySetting.QBFilePath == "")
                {
                    ClsCommon.Mode = "Config";
                    Application.Run(ClsCommon.objConfig);
                }
                else
                {
                    ClsCommon.Mode = "Sync";
                    Application.Run(ClsCommon.objSync);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Main:" + ex.Message);
            }
        }

        public static void dispose()
        {
            CommonRef.CloseQBSession();
        }

        private static void CloseInstance()
        {
            try
            {
                int i = 0;
                System.Diagnostics.Process[] oProcesses = System.Diagnostics.Process.GetProcesses(".");

                foreach (System.Diagnostics.Process oProcess in oProcesses)
                {
                    if (oProcess.ProcessName.ToLower() == "qbintegration")
                    {
                        i += 1;
                        try
                        {
                            if (i > 1)
                                oProcess.Kill();
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Application.Exit();
            }
        }

    }
}
