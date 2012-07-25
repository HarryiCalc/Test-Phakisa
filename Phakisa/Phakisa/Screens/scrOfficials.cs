using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.IO;
using Analysis = clsAnalysis;
using TB = clsTable;
using DB = clsDBase;
using Base = clsMain;
using System.Threading;
using System.Data.OleDb;
using MetaReportRuntime;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using System.Net;
using System.Net.Mail;
namespace Phakisa
{
    public partial class scrOfficials : Form
    {
        #region Declarations Server Project

        int columnnr = 0;
        int intNoOfDays = 0;
        int noOFDay = 0; 
        DateTime sheetfhs = new DateTime();
        DateTime sheetlhs = new DateTime();
        DataTable fixShifts = new DataTable();
        int intStartDay = 0;
        int intEndDay = 0;
        int intStopDay = 0;
        int workedShiftsFixedClockedShift = 0; 
        int exitValue = 0; 
        string searchEmplNr = "";
        string searchEmplName = "";
        string searchEmplGang = "";
        string searchEmplNr2 = ""; 
        string searchEmplGang2 = ""; 
        string Path = string.Empty;
        bool blTablenames = true;
        
        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        clsShared Shared = new clsShared();
        clsTableFormulas TBFormulas = new clsTableFormulas();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsAnalysis.clsAnalysis Analysis = new clsAnalysis.clsAnalysis();
        SqlConnection myConn = new SqlConnection();
        SqlConnection AConn = new SqlConnection();
        SqlConnection AAConn = new SqlConnection();
        SqlConnection BaseConn = new SqlConnection();
        SqlConnection _ADTeamConn = new SqlConnection();
        System.Collections.Hashtable buttonCollection = new System.Collections.Hashtable();
        //MetaReport.MetaReportApp tt = new MetaReport.MetaReportApp();
        //MetaReportRuntime.App uu = new MetaReportRuntime.App();
        List<string> lstTableColumns = new List<string>();

        Dictionary<string, string> dict = new Dictionary<string, string>();
        Dictionary<string, string> GangTypes = new Dictionary<string, string>();
        Dictionary<string, string> dictGridValues = new Dictionary<string, string>();
        Dictionary<string, string> dictPrimaryKeyValues = new Dictionary<string, string>();

        string strEarningsCode = string.Empty;
        string strprevPeriod = string.Empty;
        string prevDatabaseName = string.Empty;
        string strWhere = string.Empty;
        string strWhereSection = string.Empty;
        string strWherePeriod = string.Empty;
        string strActivity = string.Empty;
        string strMiningIndicator = string.Empty;
        string strMO = string.Empty;
        string strEmployeeMonthshifts = string.Empty;
        string strServerPath = string.Empty;
        string strName = string.Empty;
        string strMetaReportCode = "BSFnupmWkNxm8ZAA1ZhlOgL8fNdMdg4zhJj/j6T0vEyG9aSzk/HPwYcrjmawRGou66hBtseT7qJE+9hbEq9jces6bcGJmtz4Ih8Fic4UIw0Kt2lEffc05nFdiD2aQC0m";

        string dbPath = string.Empty;

        string[] ClockedShifts = new string[5];
        string[] OffShifts = new string[5];
        int intFiller = 0;

        List<string> lstNames = new List<string>();
        List<string> _lstColumnNames = new List<string>();

        Int64 intProcessCounter = 0;
        StringBuilder strSqlAlter = new StringBuilder();

        DataTable Survey = new DataTable();
        DataTable Labour = new DataTable();
        DataTable Miners = new DataTable();
        DataTable Designations = new DataTable();
        DataTable Participants = new DataTable();
        DataTable Clocked = new DataTable();
        DataTable Rates = new DataTable();
        DataTable EmplPen = new DataTable();
        DataTable Configs = new DataTable();
        DataTable Officials = new DataTable();
        DataTable Calendar = new DataTable();
        DataTable PayrollSends = new DataTable();
        DataTable earningsCodes = new DataTable();
        DataTable Status = new DataTable();
        DataTable BonusShifts = new DataTable();
        DataTable newDataTable = new DataTable();
        DataTable _formulas = new DataTable();
        DataTable _newDataTable = new DataTable();
        string[] arrArgs = new string[1] { "" };

        SqlDataAdapter minersTA = new SqlDataAdapter();
        BindingSource bSource = new BindingSource();
        SqlCommandBuilder _cmdBuilder = new SqlCommandBuilder();


        //**************************************************************
        //*************  PHAKISA APP.CONFIG BASIL = FS3032\SQLEXPRESS
        //     <add key="DevelopmentIntegrity" value="Trusted_Connection = True" />
        //<add key="DevelopmentServerPath" value="QwA6AFwAXABpAEMAYQBsAGMAXABcAEgAYQByAG0AbwBuAHkAXABcAFAAaABhAGsAaQBzAGEAXABcAEQAZQB2AGUAbABvAHAAbQBlAG4AdABcAFwARABhAHQAYQBiAGEAcwBlAHMAXABcAEQAYQB0AGEA" />
        //<add key="DevelopmentServerName" value="RgBTADMAMAAzADIAXABTAFEATABFAFgAUABSAEUAUwBTAA==" />
        //<add key="DevelopmentBackupPath" value="QwA6AFwAaQBDAGEAbABjAFwASABhAHIAbQBvAG4AeQBcAFAAaABhAGsAaQBzAGEAXABEAGUAdgBlAGwAbwBwAG0AZQBuAHQAXABEAGEAdABhAGIAYQBzAGUAcwA="/>
        //<add key="DevelopmentDrive" value="C:"/>

        private ExcelDataReader.ExcelDataReader spreadsheet = null;

        ToolTip tooltip = new ToolTip();
        #endregion

        public scrOfficials()
        {
            InitializeComponent(); 

        }

        internal void scrOfficialsLoad(string Period, string Region, string BussUnit, string Userid, string MiningType, string BonusType, string Environment)
        {
            #region disable all functions
            //Disable all menu functions.
            foreach (ToolStripMenuItem IT in menuStrip1.Items)
            {
                if (IT.DropDownItems.Count > 0)
                {
                    foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                    {
                        if (ITT.DropDownItems.Count > 0)
                        {
                            foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                            {
                                ITTT.Enabled = false;
                            }
                        }
                        else
                        {
                            ITT.Enabled = false;
                        }
                    }
                }
                else
                {
                    IT.Enabled = false;
                }
            }
            #endregion

            #region declarations
            BusinessLanguage.Period = Period;
            BusinessLanguage.Region = Region;
            BusinessLanguage.BussUnit = BussUnit;
            BusinessLanguage.Userid = Userid;
            BusinessLanguage.MiningType = MiningType;
            BusinessLanguage.BonusType = BonusType;
            txtMiningType.Text = MiningType;
            txtBonusType.Text = BonusType;
            strServerPath = Environment;

            BusinessLanguage.Env = Environment.Trim();
            txtDatabaseName.Text = "OFFSER2000";
            //Display dbname in text box
            //txtDatabaseName.Text = txtDatabaseName.Text.Trim() + BusinessLanguage.Period;
            Base.DBName = txtDatabaseName.Text.Trim();
            Base.Period = BusinessLanguage.Period;

            //Setup the environment BEFORE the databases are moved to the classes.  This is because the environment path forms
            //part of the fisical name of the db

            setEnvironment();

            Base.DBName = txtDatabaseName.Text.Trim();
            TB.DBName = txtDatabaseName.Text.Trim();

            #endregion

            #region Connections
            //Open Connections and create classes

            AAConn = Analysis.AnalysisConnection;
            AAConn.Open();
            BaseConn = Base.BaseConnection;
            BaseConn.Open();

            #endregion

            DataTable useraccess = Base.SelectAccessByUserid(BusinessLanguage.Userid, Base.BaseConnectionString);

            #region Assign useraccess

            //BusinessLanguage.BussUnit = useraccess.Rows[0]["BUSSUNIT"].ToString().Trim();
            BusinessLanguage.Resp = useraccess.Rows[0]["RESP"].ToString().Trim();

            foreach (DataRow dr in useraccess.Rows)
            {
                string strCodeName = dr[6].ToString().Trim();
                foreach (ToolStripMenuItem IT in menuStrip1.Items)
                {
                    if (IT.DropDownItems.Count > 0)
                    {
                        foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                        {
                            if (ITT.DropDownItems.Count > 0)
                            {
                                foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                                {
                                    if (ITTT.Name.Trim() == strCodeName)
                                    {
                                        ITTT.Enabled = true;
                                    }
                                }
                            }
                            else
                                if (ITT.Name.Trim() == strCodeName)
                                {
                                    ITT.Enabled = true;
                                }
                        }
                    }
                    else
                    {
                        if (IT.Name.Trim() == strCodeName)
                        {
                            IT.Enabled = true;
                        }

                    }
                }

            }
            #endregion

            #region General
            //Display user details
            txtUserDetails.Text = BusinessLanguage.Userid + " - " + BusinessLanguage.Region + " - " + BusinessLanguage.BussUnit;
            //txtDatabaseName.Text = BusinessLanguage.BussUnit;

            txtPeriod.Text = BusinessLanguage.Period;

            // Set up the delays for the ToolTip.
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            //Force the ToolTip text to be displayed whether or not the form is active.
            tooltip.ShowAlways = true;

            //Set up the ToolTip text for the Button and Checkbox.
            tooltip.SetToolTip(this.btnImportADTeam, "Clocked Shifts");
            tooltip.SetToolTip(this.tabLabour, "Bonus Shifts");
            tooltip.SetToolTip(this.btnRefreshSections,"This button will load all the MEASSECTIONS from clockedshifts that are not currently on Calendar");
            //tooltip.SetToolTip(this.btnSearch, "Search");

            listBox2.Enabled = false;
            listBox3.Enabled = false;


            #endregion

            #region Status button collection

            //Add the buttons needed for this bonus scheme and that are on the STATUS tab.
            buttonCollection["tabCalendar"] = btnLockCalendar;
            buttonCollection["tabLabour"] = btnLockBonusShifts;
            buttonCollection["tabOfficials"] = btnLockOfficials; 
            buttonCollection["tabEmplPen"] = btnLockEmplPen;
            buttonCollection["OFFICIALSEARN10"] = btnBaseCalcs;
            buttonCollection["OFFICIALSEARN40"] = btnOfficialsCalcs;
            buttonCollection["Input Process"] = btnInputProcess;
            buttonCollection["Bonus Report Process - Phase 1"] = btnBonusPrints;
            #endregion

            #region BaseData Extracts

            //Extract Base data
            extractDesignations();
            extractConfiguration();

            #endregion

            //Extract Tab Info
            loadInfo();
             
           
        }

        private void setEnvironment()
        {

            Base.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Base.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Base.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Base.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Base.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.MachineName = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "MachineName"].Trim();
            Base.BaseConnectionString = Base.ServerName;
            Base.Directory = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerPath"])).Trim();

            Analysis.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Analysis.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Analysis.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Analysis.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Analysis.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Analysis.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.ClockConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.DBConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.StopeConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.DevConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.BackupPath = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "BackupPath"])).Trim();
            Base.ADTeamConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.StopeDatabaseName = "STPTM2000" + BusinessLanguage.Period;
            Base.DevDatabaseName = "DEVTM2000" + BusinessLanguage.Period;

            #region oleDBConnectionStringBuilder
            OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
            builder.ConnectionString = @"Data Source=" + Base.ServerName;
            builder.Add("Provider", "SQLOLEDB.1");
            builder.Add("Initial Catalog", Base.DBName);
            //builder.Add("Persist Security Info", "False");
            builder.Add("User ID", Base.Userid);
            builder.Add("Password", Base.PWord);

            string strdb = Base.DBName;

            if (strServerPath.ToString().Contains("Development") || strServerPath.ToString().Contains("Support"))
            {
                strServerPath = "Development";

                Base.DBConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.StopeConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.BaseConnectionString = Environment.MachineName.Trim() + Base.ServerName;
                Base.ADTeamConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
            }

            //string strPath = Base.Directory.Replace("data\\", "reports\\") + strdb.Replace(BusinessLanguage.Period, "").Replace("1000", "Conn") + ".udl";
            //string strPath = "z:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            string strPath = "c:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            //MessageBox.Show("MEtatreport path en connfile :" + strPath.Trim());

            FileInfo fil = new FileInfo(strPath);

            try
            {
                File.Delete(strPath);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("delete of udl failed: " + ex.Message);
            }

            switch (strServerPath)
            {
                case "Test":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;


                case "Development":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Integrated Security", "SSPI");
                    builder.Add("Trusted_Connection", "True");
                    break;

                case "Production":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;

            }

            //MessageBox.Show("Path: " + strPath);
            bool _check = Shared.CreateUDLFile(strPath, builder);

            if (_check)
            { }
            else
            {
                MessageBox.Show("Error in creation of UDL file", "ERROR", MessageBoxButtons.OK);
            }
            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            #endregion

            myConn.ConnectionString = Base.DBConnectionString;

            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        }

        static void CreateUDLFile(string FileName, OleDbConnectionStringBuilder builder)
        {
            try
            {
                string conn = Convert.ToString(builder);
                MSDASC.DataLinksClass aaa = new MSDASC.DataLinksClass();
                aaa.WriteStringToStorage(FileName, conn, 1);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in creation of UDL file - " + ex.Message, "ERROR", MessageBoxButtons.OK);
            }
        }

        private void extractEarningsCodes()
        {

            earningsCodes = Base.SelectEarningsCodes(Base.DBConnectionString);

        }

        public void extractDBTableNames(ListBox lstbox)
        {
            connectToDB();

            if (myConn.State == ConnectionState.Open)
            {
                List<string> lstTableNames = Base.getListOfTableNamesInDatabase(Base.DBConnectionString);
                Base.DBTables = lstTableNames;
                lstbox.Items.Clear();
                listBox1.SelectionMode = SelectionMode.One;
                blTablenames = true;
                switch (lstTableNames.Count)
                {
                    case 0:
                        lstbox.Items.Add("No tables in database");
                        break;
                    default:
                        foreach (string s in lstTableNames)
                        {
                            lstbox.Items.Add(s);
                        }
                        break;
                }

            }
        }

        private void extractConfiguration()
        {

            Configs = Base.SelectConfigs(Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType);

            grdConfigs.DataSource = Configs;

        }

        private void extractDesignations()
        {
            
            cboOfficialsDesignation.Items.Clear();
            Designations = Base.GetDataByDestination("grdParticipants", Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType);

            foreach (DataRow x in Designations.Rows)
            {
                cboParticipantsDesignation.Items.Add(x["DESIGNATION"].ToString().Trim() + "  -  " + x["DESIGNATION_DESC"].ToString().Trim());
                cboOfficialsDesignation.Items.Add(x["DESIGNATION"].ToString().Trim() + "  -  " + x["DESIGNATION_DESC"].ToString().Trim());
            }
        }

        private void extractEarningsCode()
        {
            //Extract the records by miningtype, bonustype and paymethod
            DataTable t = Base.GetDataByMintypeBontypePaytype(txtMiningType.Text, txtBonusType.Text, "3", Base.BaseConnectionString);
            strEarningsCode = t.Rows[0]["EARNINGSCODE"].ToString().Trim();
        }

        private void loadInfo()
        {

            strWherePeriod = "  where period = '" + BusinessLanguage.Period + "'";
            //Check if records in calendar exists with the selected period
            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "CALENDAR", strWherePeriod);

            if (Calendar.Rows.Count > 0)
            {

                //Run the extraction of the primary keys on its own threat.
                Shared.extractPrimaryKeys(Base);

                //Run the extraction of the views.
                Shared.createViews(Base);

                if (myConn.State == ConnectionState.Open)
                {
                    evaluateAll();
                }
                else
                {
                    connectToDB();
                    evaluateAll();
                }

                //Create the tab names
                foreach (TabPage tp in tabInfo.TabPages)
                {
                    tp.Text = tp.Tag.ToString();
                }

                listBox2.SelectedIndex = 0;
                //listBox2_SelectedIndexChanged("Method", null);

            }
            else
            {
                //NO....
                //1. Get Previous months info  ==> MAG NIE MEER HIERIN GAAN NIE!!!!!!!!!!!!!!!!!!!!!!!

                getHistory();

                //2. Check if PREVIOUS months DB exists
                //if (BusinessLanguage.checkIfFileExists(Base.Directory + "\\" + prevDatabaseName + Base.DBExtention))
                //{
                //3. If exist - Create this selected DB and copy Formulas, Rates and Factors to the new database.
                DialogResult result = MessageBox.Show("Do you want to start a new Bonus Period: " + BusinessLanguage.Period + "?",
                                       "Information", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        this.Cursor = Cursors.WaitCursor;
                        backupAndRestoreDB();
                        copyFormulas();
                        extractDBTableNames(listBox1);
                        //Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period.Trim(), strprevPeriod.Trim());
                        //Base.deleteExtras2000(Base.DBConnectionString);
                        //createAndCopyCalendar();

                        //Run the extraction of the primary keys on its own threat.
                        Shared.extractPrimaryKeys(Base);
                        evaluateAll();

                        //Create the tab names
                        foreach (TabPage tp in tabInfo.TabPages)
                        {
                            tp.Text = tp.Tag.ToString();
                        }

                        listBox2.SelectedIndex = 0;
                        listBox2_SelectedIndexChanged("Method", null);
                        this.Cursor = Cursors.Arrow;
                        break;

                    case DialogResult.No:
                        btnSelect_Click("METHOD", null);
                        break;
                }
            }
        }

        private void evaluateAll()
        { 
            evaluateCalendar();
            //evaluateClockedShifts();
            evaluateLabour();
            evaluateOfficials();
            evaluateParticipants();
            evaluateEmployeePenalties();
            evaluateRates();
            extractDBTableNames(listBox1);
        }

        private void evaluateOfficials()
        {
            // Display die Miners info
            Officials.Rows.Clear();

            loadOfficials();
        }

        private void loadOfficials()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "OFFICIALS");

            if (intCount > 0)
            {
                //YES

                Officials = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Officials ", " where period = '" + BusinessLanguage.Period.Trim() + "'");

            }

            if (Officials.Rows.Count > 0)
            {
                //amp
                string strLSH = Officials.Rows[0]["LSH"].ToString().Trim();
                try
                {
                DateTime LSH = Convert.ToDateTime(strLSH);
                string Mnth = string.Empty;
                string Day = string.Empty;

                foreach (DataColumn dc in Officials.Columns)
                {
                    if (dc.Caption.Substring(0, 3) == "DAY")
                    {
                        double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                        string strTemp = Officials.Rows[0]["FSH"].ToString().Trim();
                        DateTime temp = Convert.ToDateTime(strTemp);
                        temp = temp.AddDays(d);
                        if (temp > LSH)  //remember the days start at 0
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            Officials.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                        }
                        else
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            Officials.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                        }
                    }
                }
                }
               catch
                {

                }
            }

            grdOfficials.DataSource = Officials;

            grdOfficials.Refresh();

        }

        private void copyFormulas()
        {
            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable dtBaseFormulas = Analysis.SelectAllFormulasPerDatabaseName(Base.DBCopyName + strprevPeriod.Trim(), Base.AnalysisConnectionString);
            if (dtBaseFormulas.Rows.Count > 0)
            {
                foreach (DataRow row in dtBaseFormulas.Rows)
                {
                    //Check if the receiving table already contains this formula.
                    object intCount = Analysis.countcalcbyname(Base.DBName + BusinessLanguage.Period.Trim(), row["TABLENAME"].ToString(),
                                      row["CALC_NAME"].ToString(), Base.AnalysisConnectionString);

                    if ((int)intCount > 0)
                    {
                        //rename the formula name to be inserted to NEW

                    }
                    else
                    {
                        //insert the formula.
                        Base.CopyFormulas(Base.DBName + strprevPeriod.Trim(),
                                          Base.DBName + BusinessLanguage.Period.Trim(),
                                          Analysis.AnalysisConnectionString);
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("No formulas exist on " + "\n" + "database: " + Base.DBCopyName + "\n" + "tablename: " + TB.TBCopyName +
                                "\n" + "therefor" + "\n" + "nothing will be copied", "Information", MessageBoxButtons.OK);
            }
        }


        private void confirmCopyandCreate()
        {
            listBox2.Items.Add("No sections found");

            this.Cursor = Cursors.WaitCursor;

            #region Create the new DB
            //Create the new database
            Base.createDatabase(Base.DBName, Base.ServerName);

            myConn = Base.DBConnection;
            myConn.Open();

            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createCalendarTable(Base.DBConnectionString);
            TB.createOffday(Base.DBConnectionString);
            TB.createEmployeePenalties(Base.DBConnectionString);

            //Extract Calendar again and insert into 
            DataTable calendar = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Calendar");
            grdCalendar.DataSource = calendar;

            listBox2.Items.Clear();
            listBox2.Items.Add("No sections exist yet");

            panel2.Enabled = false;
            panel3.Enabled = false;
            panel4.Enabled = false;

            this.Cursor = Cursors.Arrow;

            #endregion

        }

        private void getHistory()
        {
            #region Generate previous months db name
            //Calculate the previous months db name
            string Year = txtPeriod.Text.Trim().Substring(0, 4);
            strprevPeriod = txtPeriod.Text.Trim();

            if (txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2) == "01")
            {
                strprevPeriod = Convert.ToString(Convert.ToInt16(Year) - 1) + "12";
            }
            else
            {
                string strMonth = Convert.ToString(Convert.ToInt16(txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2)) - 1);
                if (strMonth.Length == 1)
                {
                    strMonth = "0" + strMonth;
                }

                strprevPeriod = Year + strMonth;
                prevDatabaseName = Base.DBName.Replace(txtPeriod.Text.Trim(), strprevPeriod);
            }

            Base.DBCopyName = prevDatabaseName;

            #endregion

        }

        private void createAndCopyCalendar()
        {

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");

            foreach (DataRow rr in Calendar.Rows)
            {
                rr["FSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(1)).ToString("yyyy-MM-dd");
                rr["LSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(31)).ToString("yyyy-MM-dd");          
            }

            TB.saveCalculations2(Calendar, Base.DBConnectionString, "", "CALENDAR");
            this.Cursor = Cursors.Arrow;
        }

        private void createAndCopyStatus()
        {
            getHistory();

            TB.createStatusTable(Base.DBConnectionString);
            myConn.Close();

            //create the Status datatable from the previous periods'table.
            Base.DBName = Base.DBCopyName;
            connectToDB();

            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");

            #region signoff from previous months DB and signon to this new DB

            myConn.Close();

            Base.DBName = TB.DBName;

            //Connect to the database that you want to copy from and load the tables into the listbox2.  Afterwards, change the db.dbname to the main database name.
            connectToDB();

            #endregion

            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            foreach (DataRow rr in Status.Rows)
            {
                strSQL.Append("insert into Status values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                              "','" + rr["STATUS"].ToString().Trim() + "','" + rr["LOCKED"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            Application.DoEvents();
            TB.InsertData(Base.DBConnectionString, "Update Status set status = 'N', locked = '0'");
            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");
            Application.DoEvents();
            this.Cursor = Cursors.Arrow;
        }



        private void backupAndRestoreDB()
        {
            //copy the data of the previous period to the current period.
            this.Cursor = Cursors.WaitCursor;
            Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period, strprevPeriod);

            this.Cursor = Cursors.Arrow;

        }

        private void evaluateInputProcessStatus()
        {

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", "where section = 'OFF' " +
                                                            " and category = 'Input Process'" +
                                                            " and period = '" + BusinessLanguage.Period + "'");

            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

            }
            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = 'Input Process'" +
                                      " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                btnLock.Text = "Lock";

            }

            evaluateStatus();

        }

        private void evaluateStatus()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "STATUS");

            if (intCount > 0)
            {
                //Status exists,  
                loadStatus();
            }
            else
            {
                //createAndCopyStatus();
            }

        }

        private void statusChangeButtonColors()
        {
            foreach (DataRow rr in Status.Rows)
            {
                if (rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "Exit")
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        btnRefresh.Visible = false;
                        btnx.Visible = false;

                        pictBox.Visible = false;
                        pictBox2.Visible = false;
                        calcTime.Enabled = false;
                    }
                }
                else
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        string strButtonName = rr["PROCESS"].ToString().Trim();
                        Control c = (Control)buttonCollection[strButtonName];
                        c.BackColor = Color.LightGreen;

                    }
                    else
                    {
                        if (rr["STATUS"].ToString().Trim() == "P")
                        {
                            string strButtonName = rr["PROCESS"].ToString().Trim();
                            Control c = (Control)buttonCollection[strButtonName];
                            c.BackColor = Color.Orange;
                        }
                        else
                        {
                            if (rr["STATUS"].ToString().Trim() == "N" &&
                                pictBox.Visible == true &&
                                rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "CALC")
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.Orange;
                            }
                            else
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.PowderBlue;
                            }
                        }
                    }
                }

                Application.DoEvents();
            }
        }
      

        private void evaluateLabour()
        {
            //xxxxxxxxxxxxxxxxxx
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

            if (intCount > 0)
            {    
                DataTable tempoff = new DataTable();

                if (Calendar.Rows.Count > 0)
                {
                    IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                                  where locks.Field<string>("SECTION").Trim() == "OFF"
                                                  where locks.Field<string>("PERIOD").Trim() == BusinessLanguage.Period.Trim()
                                                  select locks;


                    tempoff = query1.CopyToDataTable<DataRow>();
                }
                else
                {
                    MessageBox.Show("No records in Calendar.  Contact support.", "ERROR", MessageBoxButtons.OK);
                }

                string strSQL = "select * from BONUSSHIFTS where period = '" + BusinessLanguage.Period + "'" ;


                Labour = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                #region Change column names
                if (Labour.Rows.Count > 0)
                {
                    //amp
                    string strLSH = tempoff.Rows[0]["LSH"].ToString().Trim();
                    DateTime LSH = Convert.ToDateTime(strLSH);
                    string Mnth = string.Empty;
                    string Day = string.Empty;

                    foreach (DataColumn dc in Labour.Columns)
                    {
                        if (dc.Caption.Substring(0, 3) == "DAY")
                        {
                            double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                            string strTemp = tempoff.Rows[0]["FSH"].ToString().Trim();
                            DateTime temp = Convert.ToDateTime(strTemp);
                            temp = temp.AddDays(d);
                            if (temp > LSH)  //remember the days start at 0
                            {
                                if (Convert.ToString(temp.Day).Length < 2)
                                {
                                    Day = "0" + Convert.ToString(temp.Day);
                                }
                                else
                                {
                                    Day = Convert.ToString(temp.Day);
                                }
                                if (Convert.ToString(temp.Month).Length < 2)
                                {
                                    Mnth = "0" + Convert.ToString(temp.Month);
                                }
                                else
                                {
                                    Mnth = Convert.ToString(temp.Month);
                                }
                                Labour.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                            }
                            else
                            {
                                if (Convert.ToString(temp.Day).Length < 2)
                                {
                                    Day = "0" + Convert.ToString(temp.Day);
                                }
                                else
                                {
                                    Day = Convert.ToString(temp.Day);
                                }
                                if (Convert.ToString(temp.Month).Length < 2)
                                {
                                    Mnth = "0" + Convert.ToString(temp.Month);
                                }
                                else
                                {
                                    Mnth = Convert.ToString(temp.Month);
                                }
                                Labour.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                            }
                        }
                    }
                }
                //amp}
                grdLabour.DataSource = Labour;
                #endregion

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_NO");
                cboNames.Items.Clear();
                cboEmplPenEmployeeNo.Items.Clear();

                foreach (string s in lstNames)
                {

                    cboEmplPenEmployeeNo.Items.Add(s.Trim());
                    cboNames.Items.Add(s.Trim());

                }

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_NAME");

                cboMinersEmpName.Items.Clear();

                foreach (string s in lstNames)
                {

                    cboMinersEmpName.Items.Add(s.Trim());

                }

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "GANG");   
                cboBonusShiftsGang.Items.Clear();


                foreach (string s in lstNames)
                {

                   
                    cboBonusShiftsGang.Items.Add(s.Trim());

                }     

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "WAGECODE");   
                cboBonusShiftsWageCode.Items.Clear();
                foreach (string s in lstNames)
                {

                    cboBonusShiftsWageCode.Items.Add(s.Trim()); 

                }    

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "LINERESPCODE");   
                cboBonusShiftsResponseCode.Items.Clear();
                foreach (string s in lstNames)
                {

                    cboBonusShiftsResponseCode.Items.Add(s.Trim());

                }     
            }

            else
            {
                MessageBox.Show("Bonus Shifts data does not exist.  Please import the data before trying to process.", "Information", MessageBoxButtons.OK);
            }

            hideColumnsOfGrid("grdLabour");
        }

        private void hideColumnsOfGrid(string gridname)
        {

            switch (gridname)
            {
                

                case "grdPayroll":
                    if (grdPayroll.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdPayroll.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdPayroll.Columns.Contains("MININGTYPE"))
                    {
                        this.grdPayroll.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdPayroll.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdPayroll.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;

                

                case "grdLabour":

                    if (grdLabour.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdLabour.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("MININGTYPE"))
                    {
                        this.grdLabour.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdLabour.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdRates":
                    if (grdRates.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdRates.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("MININGTYPE"))
                    {
                        this.grdRates.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdRates.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdCalendar":
                    if (grdCalendar.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdCalendar.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("MININGTYPE"))
                    {
                        this.grdCalendar.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdCalendar.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
            }
        }

        private void evaluateCalendar()
        {
            panel3.Enabled = true;
            panel4.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true;

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "CALENDAR");

            if (intCount > 0)
            {
                //Calendar exists,
                loadCalendar();
                loadDatePickers(0);
                loadSectionsFromCalendar();
            }
            else
            {
                createAndCopyCalendar();
            }
        }

        private void loadCalendar()
        {
            // Display die calendar info

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", " where period = '" + BusinessLanguage.Period + "'");

            grdCalendar.DataSource = Calendar;


        }

        private void loadStatus()
        {
            // Display die STATUS info

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status",strWhere);

            if (Status.Rows.Count > 0)
            {
                statusChangeButtonColors();
            }
            else
            {
                //createAndCopyStatus();
            }
        }

        private void loadDatePickers(int Position)
        {

            #region check number of shifts

            if (Calendar.Rows.Count > 0)
            {
                DateTime EarliestStart = Convert.ToDateTime(Calendar.Rows[0]["FSH"].ToString().Trim());
                DateTime LatestEnd = Convert.ToDateTime(Calendar.Rows[0]["LSH"].ToString().Trim());

                for (int i = 0; i <= Calendar.Rows.Count - 1; i++)
                {

                    if (Convert.ToDateTime(Calendar.Rows[i]["FSH"].ToString().Trim()) < EarliestStart)
                    {
                        EarliestStart = Convert.ToDateTime(Calendar.Rows[i]["FSH"].ToString().Trim());
                    }
                    if (Convert.ToDateTime(Calendar.Rows[i]["LSH"].ToString().Trim()) > LatestEnd)
                    {
                        LatestEnd = Convert.ToDateTime(Calendar.Rows[i]["LSH"].ToString().Trim());
                    }

                }

                dateTimePicker1.Value = EarliestStart;
                dateTimePicker2.Value = LatestEnd;
                intNoOfDays = Base.calcNoOfDays(LatestEnd, EarliestStart);

                if (intNoOfDays > 45)
                {
                    MessageBox.Show("Your earliest date is: " + EarliestStart + " and your latest date is: " + LatestEnd + Environment.NewLine +
                        " that is " + intNoOfDays + " days."  + Environment.NewLine + " The difference between earliest start date and latest end date should" + 
                        " be 45 days or less. " + 
                        Environment.NewLine +
                        " Please set your dates on calendar accordingly.", "ERROR", MessageBoxButtons.OK);
                }

            }

            #endregion

        }

        private void loadSectionsFromCalendar()
        {
            lstNames.Clear();
            lstNames.Add("OFF");

            if (lstNames.Count > 0)
            {
                //xxxxxxxxxx
                txtSelectedSection.Text = "OFF";
                label15.Text = "OFF";
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "' and period = '" + BusinessLanguage.Period + "'";
                strWhereSection = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "'";
                listBox2.Items.Clear();

                if (lstNames.Count > 1)
                {
                    foreach (string s in lstNames)
                    {
                        if (s != "XXX")
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
                else
                {
                    if (lstNames.Count == 1)
                    {
                        foreach (string s in lstNames)
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
            }
        }

        private void evaluatePayroll()
        {
            // Display die Ganglink info
            PayrollSends.Rows.Clear();

            loadPayroll();

        }

        private void loadPayroll()
        {
            //Check if Payroll exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");

            if (intCount > 0)
            {
                //YES
                PayrollSends = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Payroll", strWhere); 
            }

            grdPayroll.DataSource = PayrollSends;    
        }

        private void loadMO()
        {
            strMO = "";
            foreach (DataRow dr in Configs.Rows)
            {
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING" && dr["PARM1"].ToString().Trim() == "MO" && dr["PARM2"].ToString().Trim() == txtSelectedSection.Text)
                {
                    for (int i = 3; i <= 5; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMO = "'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strMO = "(" + strMO.Trim() + ")";
                }
            }
        }

        private void evaluateParticipants()
        {
            // Display die Abnormal info
            Participants.Rows.Clear();

            loadParticipants();

        }

        private void loadParticipants()
        {
            //Check if Participants exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Participants");

            if (intCount > 0)
            {
                //YES

                Participants = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Participants",  " where period = '" + BusinessLanguage.Period.Trim() + "'");

                lstNames = TB.loadDistinctValuesFromColumn(Participants, "HOD");

                foreach (string s in lstNames)
                {

                    cboParticipantsHOD.Items.Add(s.Trim());
                }

                lstNames = TB.loadDistinctValuesFromColumn(Participants, "HOD");

                foreach (string s in lstNames)
                {

                    cboParticipantsHOD.Items.Add(s.Trim());
                }

            }
            else
            {
                //NO - Participants DOES NOT EXIST 
            }

            grdParticipants.DataSource = Participants;

            grdParticipants.Refresh();
        }

        private void evaluateRates()
        {
            // Display die Abnormal info
            Rates.Rows.Clear();

            loadRates();

        }

        private void evaluateEmployeePenalties()
        {
            // Display die EmployeePenalties info
            EmplPen.Rows.Clear();

            loadEmployeePenalties();

        }

        private void loadEmployeePenalties()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            if (intCount > 0)
            {
                //YES

                EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES",  " where period = '" + BusinessLanguage.Period.Trim() + "'");


            }
            else
            {
                //NO
                //Check if Bonusshifts Exists

                intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

                if (intCount > 0)
                {
                    TB.createEmployeePenalties(Base.DBConnectionString);
                    TB.TBName = "EMPLOYEEPENALTIES";
                    EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES ", strWhere);

                }
                else
                {
                }

            }

            grdEmplPen.DataSource = EmplPen;

            grdEmplPen.Refresh();

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            Application.Exit();

            //**********Old code

            //DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

            //switch (result)
            //{
            //    case DialogResult.Yes:
            //        this.Close();
            //        //scrMain main = new scrMain();
            //        //main.MainLoad(BusinessLanguage, DB, Survey, Labour, Miners, Designations, Occupations, Clocked, EmplList, EmplPen, Configs);
            //        //main.ShowDialog();
            //        myConn.Close();
            //        AAConn.Close();
            //        AConn.Close();
            //        this.Close();
            //        break;

            //    case DialogResult.No:
            //        break;
            //}

        }

        private void connectToDB()
        {

            if (myConn.State == ConnectionState.Closed)
            {
                try
                {
                    myConn.Open();
                }
                catch (SystemException eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }
        }

        private void loadRates()
        {
            //Check if ABNORMAL exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Rates");

            if (intCount > 0)
            {
                //YES

                Rates = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Rates",strWherePeriod);

            }
            else
            {
                //NO - Rates DOES NOT EXIST 
            }

            grdRates.DataSource = Rates;

            grdRates.Refresh();

        }

        
       

       

        private void importTheSheet(string importFilename)
        {
            string path = BusinessLanguage.InputDirectory + Base.DBName;

            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                string filename = BusinessLanguage.InputDirectory + Base.DBName + importFilename;
                bool fileCheck = BusinessLanguage.checkIfFileExists(filename);

                if (fileCheck)
                {
                    FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                    spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                    fs.Close();
                    //If the file was SURVEY, all sections production data will be on this datatable.
                    //Only the selected section's data must be saved.

                    saveTheSpreadSheetToTheDatabase();
                }
                else
                {
                    MessageBox.Show("File " + filename + " - does not exist", "Check", MessageBoxButtons.OK);
                }

                //Check if file exists
                //If not  = Message
                //If exists ==>  Import
            }
            catch
            {
                MessageBox.Show("File " + importFilename + " - is inuse by another package?", "Check", MessageBoxButtons.OK);
            }
        }

        private void saveTheSpreadSheetToTheDatabase()
        {
            foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
            {
                if (dt.TableName == "SURVEY" || dt.TableName == "Survey")
                {
                    for (int i = 1; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][3].ToString().Trim() == txtSelectedSection.Text.Trim())
                        {
                        }
                        else
                        {
                            dt.Rows[i].Delete();

                        }
                    }

                }

                dt.AcceptChanges();
                //checker = true;

                TB.TBName = dt.TableName.ToString().ToUpper();
                TB.recreateDataTable();

                //Extract column names
                string strColumnHeadings = TB.getFirstRowValues(dt, Base.AnalysisConnectionString);

                switch (strColumnHeadings)
                {
                    case null:
                        break;

                    case "":
                        break;

                    default:


                        if (myConn.State == ConnectionState.Closed)
                        {
                            try
                            {
                                myConn = Base.DBConnection;
                                myConn.Open();

                                //create a table
                                bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);

                                if (tableCreate)
                                {
                                    MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    MessageBox.Show("Try again after correction of spreadsheet - input data.", "Information", MessageBoxButtons.OK);
                                }

                                //checker = false;
                            }
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.GetHashCode() + " " + ex.ToString(), "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //create a table
                            bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                            if (tableCreate)
                            {
                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);
                                MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);

                            }
                            else
                            {
                                MessageBox.Show("Data was not imported.", "Information", MessageBoxButtons.OK);
                            }
                        }

                        break;
                }
            }
        }
       
        private String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" + excelFile + ";Extended Properties=Excel 12.0;";
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch(Exception EX) 
            {
                MessageBox.Show("Error on import: " + EX.Message);
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }
        private void btnImportADTeam_Click(object sender, EventArgs e)
        {
            DataTable temp = new DataTable();

            int intCalendarProcesses = checkLockCalendarProcesses();

            if (intCalendarProcesses > 0)
            {
                MessageBox.Show("Please finalize Calendar before importing your shifts.");
            }
            else
            {
                if (Labour.Rows.Count > 0)
                {
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").Trim() == "tabLabour"
                                                  where locks.Field<string>("SECTION").Trim() == "OFF"
                                                  where locks.Field<string>("PERIOD").Trim() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        loadDatePickers(0);
                        if (intNoOfDays <= 45)
                        {
                            refreshLabour();
                        }
                        else
                        {
                            MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                MessageBoxButtons.OK);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("No records on Status for the Section,Period and tabLabour");
                    }
                }
                else
                {
                    evaluateStatus();
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").Trim() == "tabLabour"
                                                  where locks.Field<string>("SECTION").Trim() == "OFF"
                                                  where locks.Field<string>("PERIOD").Trim() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        if (temp.Rows.Count > 0)
                        {
                            if (temp.Rows[0]["STATUS"].ToString().Trim() == "N")
                            {

                                loadDatePickers(0);
                                if (intNoOfDays <= 45)
                                {
                                    refreshLabour();
                                }
                                else
                                {
                                    MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                        MessageBoxButtons.OK);
                                }

                            }
                            else
                            {
                                MessageBox.Show("BonusShifts is locked. Please unlock before refresh.  You WILL loose all previous updates.",
                                    "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not find LOCK records on STATUS for selected period.  Please re-select the section.", "Information", MessageBoxButtons.OK);
                    }
                }

            }
        }

        private void refreshLabour2()
        {
            #region extract the  FSH from the database
            this.Cursor = Cursors.WaitCursor;
            //This is the refresh from the ADTeam database.
            _ADTeamConn = Base.ADTeamConnection;
            _ADTeamConn.Open();

            DataTable ADTeam = TB.createDataTableWithAdapter(Base.ADTeamConnectionString, "select TOP 1 *  from FREEGOLD_EMPLOYEEDETAIL");

            DateTime _lastRunDate = Convert.ToDateTime(ADTeam.Rows[0]["lastrundate"]);

            int intStart = Base.calcNoOfDays(_lastRunDate, dateTimePicker1.Value) + 1;
            int intEnd = intStart - 44;

            if (intEnd <= 0)
            {
                intEnd = 1;
            }

            if (intStart > 100)
            {
                intStart = 100;
            }

            int intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);

            DataTable dt = TB.ExtractADTeamShifts(Base.ADTeamConnectionString, intStart, intEnd, dateTimePicker1.Value,
                                                  dateTimePicker2.Value, intStart, intEnd,
                                                  BusinessLanguage.Period, txtSelectedSection.Text.Trim(), BusinessLanguage.MiningType,
                                                  BusinessLanguage.BonusType, BusinessLanguage.BussUnit, " where bussunit = 'JJ'");

            foreach (DataRow row in dt.Rows)
            {

                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }

            TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");


            //string tst = string.Empty;
            //for (int i = 0; i <= dt.Columns.Count - 1; i++)
            //{
            //    tst = tst.Trim() + "-" + dt.Columns[i].ColumnName.Trim();
            //}

            //MessageBox.Show(intStart.ToString().Trim() + "-" + intEnd.ToString().Trim() + "-" + intNoOfDays.ToString().Trim() + "-" + tst);

            #endregion

            #region Apply offdays
            if (dt.Rows.Count > 0)
            {
                Clocked = dt.Copy();
                //Update clockedshifts with offday calendar data
                UpdateClockedShifts();
                dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                Application.DoEvents();

                //grdClocked.DataSource = dt;
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = " Select distinct t1.*,'0' as SHIFTS_WORKED, " +
                                " '0' as AWOP_SHIFTS,'0' as Q_SHIFTS,t2.MEASSECTION,t2.FSH as FSH_Participant,t2.LSH as LSH_Participant  " +
                                " from Clockedshifts AS T1, PARTICIPANTS AS T2 where T1.section = 'OFF' " +
                                " AND T1.SECTION = T2.SECTION " +
                                " AND T1.EMPLOYEE_NO = T2.EMPLOYEE_NO " +
                                " AND T1.PERIOD = T2.PERIOD;  ";

                DataTable noDups = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                BonusShifts = removeDuplicateRecords(noDups); 

                #region count the shifts
                //Count the the shifts

                DialogResult result = MessageBox.Show(@"Do you want to REPLACE the current BONUSSHIFTS for section " +
                                                      txtSelectedSection.Text.Trim() + @" ?", @"QUESTION", MessageBoxButtons.OKCancel);

                switch (result)
                {
                    case DialogResult.OK:

                        #region Calculate the shifts per employee en output to bonusshifts
                        //Then count each employee's shifts individually 

                        foreach (DataRow dr in BonusShifts.Rows)
                        {
                            DateTime drFSH = Convert.ToDateTime(dr["FSH"].ToString());
                            DateTime drLSH = Convert.ToDateTime(dr["LSH"].ToString());
                            DateTime drFSH_Participant = Convert.ToDateTime(dr["FSH_Participant"].ToString());
                            DateTime drLSH_Participant = Convert.ToDateTime(dr["LSH_Participant"].ToString());

                            intStart = Base.calcNoOfDays(drFSH_Participant,drFSH );
                            intEnd = Base.calcNoOfDays(drLSH_Participant,drLSH);

                            if (intStart <= 0)
                            {
                                intStart = 0;
                            }

                            if (intEnd <= 0)
                            {
                                intEnd = 44 + intEnd;
                            }
                            else
                            {
                                intEnd = 44;
                            }

                            extractAndCalcShifts(intStart, intEnd, dr);

                        }

                        #endregion


                        //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                        DataColumn dcPeriod = new DataColumn();
                        dcPeriod.ColumnName = "PERIOD";
                        BonusShifts.Columns.Remove("PERIOD");
                        BonusShifts.AcceptChanges();
                        InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

                        foreach (DataRow dr in BonusShifts.Rows)
                        {
                            dr["PERIOD"] = BusinessLanguage.Period;
                        }

                        BonusShifts.Columns.Remove("FSH_Participant");
                        BonusShifts.Columns.Remove("LSH_Participant");
                        BonusShifts.AcceptChanges();

                        string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                                           "' and period = '" + BusinessLanguage.Period.Trim() + "'";

                        TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");


                        // Update the linerespcode on bonusshifts to the meassection on participants
                        TB.InsertData(Base.DBConnectionString, "UPDATE Bonusshifts SET Bonusshifts.linerespcode = Participants.meassection " +
                                      " FROM Bonusshifts  INNER JOIN  Participants ON Bonusshifts.Section = Participants.Section " +
                                      " and Bonusshifts.Period =  Participants.PERIOD and Bonusshifts.Period = '" + BusinessLanguage.Period.Trim() + 
                                      "' and Bonusshifts.employee_no = Participants.employee_no");


                        break;

                    case DialogResult.Cancel:
                        break;

                }

                #endregion

                #endregion

                this.Cursor = Cursors.Arrow;


            }
            else
            {
                MessageBox.Show(@"No ADTeam records available for dates on calendar and/or businessunit parameters", "Information", MessageBoxButtons.OK);
            }
        }

        private void recreateBonusshifts()
        {
            #region Calculate the shifts per employee en output to bonusshifts

            string strSQL = " Select distinct t1.*,'0' as SHIFTS_WORKED, " +
                            " '0' as AWOP_SHIFTS,'0' as Q_SHIFTS,t2.MEASSECTION,t2.FSH as FSH_Participant,t2.LSH as LSH_Participant  " +
                            " from Clockedshifts AS T1, PARTICIPANTS AS T2 where T1.section = 'OFF' " +
                            " AND T1.SECTION = T2.SECTION " +
                            " AND T1.EMPLOYEE_NO = T2.EMPLOYEE_NO " +
                            " AND T1.PERIOD = T2.PERIOD;  ";

            DataTable noDups = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            BonusShifts = removeDuplicateRecords(noDups);

            #region count the shifts
            //Count the the shifts


                    #region Calculate the shifts per employee en output to bonusshifts
                    //Then count each employee's shifts individually 

                    foreach (DataRow dr in BonusShifts.Rows)
                    {
                        DateTime drFSH = Convert.ToDateTime(dr["FSH"].ToString());
                        DateTime drLSH = Convert.ToDateTime(dr["LSH"].ToString());
                        DateTime drFSH_Participant = Convert.ToDateTime(dr["FSH_Participant"].ToString());
                        DateTime drLSH_Participant = Convert.ToDateTime(dr["LSH_Participant"].ToString());

                        int intStart = Base.calcNoOfDays(drFSH_Participant, drFSH);
                        int intEnd = Base.calcNoOfDays(drLSH_Participant, drLSH);

                        if (intStart <= 0)
                        {
                            intStart = 0;
                        }

                        if (intEnd <= 0)
                        {
                            intEnd = 44 + intEnd;
                        }
                        else
                        {
                            intEnd = 44;
                        }

                        extractAndCalcShifts(intStart, intEnd, dr);

                    }

                    #endregion


                    //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                    DataColumn dcPeriod = new DataColumn();
                    dcPeriod.ColumnName = "PERIOD";
                    BonusShifts.Columns.Remove("PERIOD");
                    BonusShifts.AcceptChanges();
                    InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

                    foreach (DataRow dr in BonusShifts.Rows)
                    {
                        dr["PERIOD"] = BusinessLanguage.Period;
                    }

                    BonusShifts.Columns.Remove("FSH_Participant");
                    BonusShifts.Columns.Remove("LSH_Participant");
                    BonusShifts.AcceptChanges();

                    string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                                       "' and period = '" + BusinessLanguage.Period.Trim() + "'";

                    TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");


                    // Update the linerespcode on bonusshifts to the meassection on participants
                    TB.InsertData(Base.DBConnectionString, "UPDATE Bonusshifts SET Bonusshifts.linerespcode = Participants.meassection " +
                                  " FROM Bonusshifts  INNER JOIN  Participants ON Bonusshifts.Section = Participants.Section " +
                                  " and Bonusshifts.Period =  Participants.PERIOD and Bonusshifts.Period = '" + BusinessLanguage.Period.Trim() +
                                  "' and Bonusshifts.employee_no = Participants.employee_no");


            #endregion

            #endregion




        }

        private void refreshLabour3()
        {

            #region extract the sheet name and FSH and LSH of the extract
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "BONTS2011");
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {

                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "BONTS2011");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
            }

            if (FilePath.Trim().Length == 0)
            {
                MessageBox.Show("ADTeam spreadsheet does not exist", "Information", MessageBoxButtons.OK);
                this.Cursor = Cursors.Arrow;
            }
            else
            {
                //excel.GetExcelSheets();
                string[] sheetNames = GetExcelSheetNames(FilePath);
                string sheetName = sheetNames[0];
            #endregion

                #region import Clockshifts
                this.Cursor = Cursors.WaitCursor;
                DataTable dt = new DataTable();

                OleDbConnection con = new OleDbConnection();
                OleDbDataAdapter da;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                        + FilePath + ";Extended Properties='Excel 12.0;'";

                /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
                * "HDR=No;" indicates the opposite.
                * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
                * Note that this option might affect excel sheet write access negative.
                */

                da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
                da.Fill(dt);

                IEnumerable<DataRow> query1 = from locks in dt.AsEnumerable()
                                              where locks.Field<string>("WAGE CODE").TrimEnd().Length > 0 &&
                                                    locks.Field<string>("WAGE CODE").TrimEnd().Substring(0, 1) != "0"
                                              select locks;

                //Temp will contain a list of the gangs for the section
                DataTable Tempdt = query1.CopyToDataTable<DataRow>();

                dt = Tempdt.Copy();

                #region remove invalid records

                //extract the column names with length less than 3.  These columns must be deleted.
                string[] columnNames = new String[dt.Columns.Count];

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (dt.Columns[i].ColumnName.Length <= 2)
                    {
                        columnNames[i] = dt.Columns[i].ColumnName;
                    }
                }

                for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
                {
                    if (string.IsNullOrEmpty(columnNames[i]))
                    {

                    }
                    else
                    {
                        dt.Columns.Remove(columnNames[i].ToString().Trim());
                        dt.AcceptChanges();
                    }
                }

                if (dt.Columns.Contains("INDUSTRY NUMBER"))
                {
                    dt.Columns.Remove("INDUSTRY NUMBER");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 1"))
                {
                    dt.Columns.Remove("BONUS 1");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 2"))
                {
                    dt.Columns.Remove("BONUS 2");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 3"))
                {
                    dt.Columns.Remove("BONUS 3");
                    dt.AcceptChanges();
                }


                #endregion

                string strSheetFSH = string.Empty;
                string strSheetLSH = string.Empty;

                //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
                string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
                string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

                //Correct the dates and calculate the number of days extracted.
                if (strSheetFSHx.Substring(6, 1) == "-")
                {
                    strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
                }
                else
                {
                    strSheetFSH = strSheetFSHx;
                }

                if (strSheetLSHx.Substring(6, 1) == "-")
                {
                    strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
                }
                else
                {
                    strSheetLSH = strSheetLSHx;
                }

                DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
                DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

                TB.InsertData(Base.DBConnectionString, "Update calendar set FSH = '" + strSheetFSH.ToString().Trim() +
                                                      "', LSH = '" + strSheetLSH.ToString() + "' where section = 'OFF'");

                //If the intNoOfDays < 44 then the days up to 44 must be filled with '-'
                int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

                if (intNoOfDays <= 44)
                {
                    for (int j = intNoOfDays + 1; j <= 44; j++)
                    {
                        dt.Columns.Add("DAY" + j);
                    }
                }
                else
                {

                }

                #region Change the column names
                //Change the column names to the correct column names.
                Dictionary<string, string> dictNames = new Dictionary<string, string>();
                DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                     "Select * from varnames");
                dictNames.Clear();

                dictNames = TB.loadDict(varNames, dictNames);
                int counter = 0;

                //If it is a column with a date as a name.
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Substring(0, 1) == "2")
                    {
                        if (counter == 0)
                        {
                            strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            if (column.Ordinal == dt.Columns.Count - 1)
                            {

                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;

                            }
                            else
                            {
                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;
                            }
                        }


                    }
                    else
                    {
                        if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                        {
                            column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                        }

                    }
                }

                dt.Columns.Add("FSH");
                dt.Columns.Add("LSH");
                dt.Columns.Add("SECTION");
                dt.Columns.Add("EMPLOYEETYPE");
                dt.Columns.Add("PERIOD");      //xxxxxxxx
                dt.Columns.Add("BUSSUNIT");
                dt.AcceptChanges();


                foreach (DataRow row in dt.Rows)
                {

                    row["FSH"] = strSheetFSH;
                    row["LSH"] = strSheetLSH;
                    row["MININGTYPE"] = "OFFICIALS";
                    row["BONUSTYPE"] = "SERVICES";
                    row["PERIOD"] = BusinessLanguage.Period;   //xxx
                    row["BUSSUNIT"] = BusinessLanguage.BussUnit;
                    row["SECTION"] = "OFF";

                    if (row["WAGECODE"].ToString().Trim() == "")
                    {
                        row["WAGECODE"] = "00000";
                    }
                    else
                    {
                    }
                    row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                        {
                            row[i] = "-";
                        }
                    }
                }

                #endregion
                //exportToExcel("c:\\", dt);
                //Write to the database
                dt.Columns["BUSSUNIT"].SetOrdinal(11);
                TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");

                Application.DoEvents();

                #region Apply offdays

                if (dt.Rows.Count > 0)
                {
                    Clocked = dt.Copy();
                    //Update clockedshifts with offday calendar data
                    UpdateClockedShifts();
                    dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                    Application.DoEvents();

                    //grdClocked.DataSource = dt;
                }
                #endregion

                #endregion



                #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = " Select distinct t1.*,'0' as SHIFTS_WORKED, " +
                                " '0' as AWOP_SHIFTS,'0' as Q_SHIFTS,t2.MEASSECTION,t2.FSH as FSH_Participant,t2.LSH as LSH_Participant  " +
                                " from Clockedshifts AS T1, PARTICIPANTS AS T2 where T1.section = 'OFF' " +
                                " AND T1.SECTION = T2.SECTION " +
                                " AND T1.EMPLOYEE_NO = T2.EMPLOYEE_NO " +
                                " AND T1.PERIOD = T2.PERIOD;  ";

                DataTable noDups = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                BonusShifts = removeDuplicateRecords(noDups);

                #region count the shifts
                //Count the the shifts

                DialogResult result = MessageBox.Show(@"Do you want to REPLACE the current BONUSSHIFTS for section " +
                                                      txtSelectedSection.Text.Trim() + @" ?", @"QUESTION", MessageBoxButtons.OKCancel);

                switch (result)
                {
                    case DialogResult.OK:

                        #region Calculate the shifts per employee en output to bonusshifts
                        //Then count each employee's shifts individually 

                        foreach (DataRow dr in BonusShifts.Rows)
                        {
                            DateTime drFSH = Convert.ToDateTime(dr["FSH"].ToString());
                            DateTime drLSH = Convert.ToDateTime(dr["LSH"].ToString());
                            DateTime drFSH_Participant = Convert.ToDateTime(dr["FSH_Participant"].ToString());
                            DateTime drLSH_Participant = Convert.ToDateTime(dr["LSH_Participant"].ToString());

                            int intStart = Base.calcNoOfDays(drFSH_Participant, drFSH);
                            int intEnd = Base.calcNoOfDays(drLSH_Participant, drLSH);

                            if (intStart <= 0)
                            {
                                intStart = 0;
                            }

                            if (intEnd <= 0)
                            {
                                intEnd = 44 + intEnd;
                            }
                            else
                            {
                                intEnd = 44;
                            }

                            extractAndCalcShifts(intStart, intEnd, dr);

                        }

                        #endregion


                        //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                        DataColumn dcPeriod = new DataColumn();
                        dcPeriod.ColumnName = "PERIOD";
                        BonusShifts.Columns.Remove("PERIOD");
                        BonusShifts.AcceptChanges();
                        InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

                        foreach (DataRow dr in BonusShifts.Rows)
                        {
                            dr["PERIOD"] = BusinessLanguage.Period;
                        }

                        BonusShifts.Columns.Remove("FSH_Participant");
                        BonusShifts.Columns.Remove("LSH_Participant");
                        BonusShifts.AcceptChanges();

                        string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                                           "' and period = '" + BusinessLanguage.Period.Trim() + "'";

                        TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");


                        // Update the linerespcode on bonusshifts to the meassection on participants
                        TB.InsertData(Base.DBConnectionString, "UPDATE Bonusshifts SET Bonusshifts.linerespcode = Participants.meassection " +
                                      " FROM Bonusshifts  INNER JOIN  Participants ON Bonusshifts.Section = Participants.Section " +
                                      " and Bonusshifts.Period =  Participants.PERIOD and Bonusshifts.Period = '" + BusinessLanguage.Period.Trim() +
                                      "' and Bonusshifts.employee_no = Participants.employee_no");


                        break;

                    case DialogResult.Cancel:
                        break;

                }

                #endregion

                #endregion

                this.Cursor = Cursors.Arrow;
            }
        }

        private void refreshLabour()
        {

            #region extract the sheet name and FSH and LSH of the extract
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "BONTS2011");
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {

                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "BONTS2011");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
            }

            if (FilePath.Trim().Length == 0)
            {
                MessageBox.Show("ADTeam spreadsheet does not exist", "Information", MessageBoxButtons.OK);
                this.Cursor = Cursors.Arrow;
            }
            else
            {
                //excel.GetExcelSheets();
                string[] sheetNames = GetExcelSheetNames(FilePath);
                string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
                this.Cursor = Cursors.WaitCursor;
                DataTable dt = new DataTable();

                OleDbConnection con = new OleDbConnection();
                OleDbDataAdapter da;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                        + FilePath + ";Extended Properties='Excel 12.0;'";

                /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
                * "HDR=No;" indicates the opposite.
                * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
                * Note that this option might affect excel sheet write access negative.
                */

                da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
                da.Fill(dt);

                IEnumerable<DataRow> query1 = from locks in dt.AsEnumerable()
                                              where locks.Field<string>("WAGE CODE").TrimEnd().Length > 0 &&
                                                    locks.Field<string>("WAGE CODE").TrimEnd().Substring(0, 1) != "0"
                                              select locks;

                //Temp will contain a list of the gangs for the section
                DataTable Tempdt = query1.CopyToDataTable<DataRow>();

                dt = Tempdt.Copy();

                #region remove invalid records

                //extract the column names with length less than 3.  These columns must be deleted.
                string[] columnNames = new String[dt.Columns.Count];

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (dt.Columns[i].ColumnName.Length <= 2)
                    {
                        columnNames[i] = dt.Columns[i].ColumnName;
                    }
                }

                for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
                {
                    if (string.IsNullOrEmpty(columnNames[i]))
                    {

                    }
                    else
                    {
                        dt.Columns.Remove(columnNames[i].ToString().Trim());
                        dt.AcceptChanges();
                    }
                }

                if (dt.Columns.Contains("INDUSTRY NUMBER"))
                {
                    dt.Columns.Remove("INDUSTRY NUMBER");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 1"))
                {
                    dt.Columns.Remove("BONUS 1");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 2"))
                {
                    dt.Columns.Remove("BONUS 2");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 3"))
                {
                    dt.Columns.Remove("BONUS 3");
                    dt.AcceptChanges();
                }

                 
                #endregion

                string strSheetFSH = string.Empty;
                string strSheetLSH = string.Empty;

                //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
                string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
                string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

                //Correct the dates and calculate the number of days extracted.
                if (strSheetFSHx.Substring(6, 1) == "-")
                {
                    strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
                }
                else
                {
                    strSheetFSH = strSheetFSHx;
                }

                if (strSheetLSHx.Substring(6, 1) == "-")
                {
                    strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
                }
                else
                {
                    strSheetLSH = strSheetLSHx;
                }

                DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
                DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

                TB.InsertData(Base.DBConnectionString, "Update calendar set FSH = '" + strSheetFSH.ToString().Trim() +
                                                      "', LSH = '" + strSheetLSH.ToString() + "' where section = 'OFF'");

                //If the intNoOfDays < 44 then the days up to 44 must be filled with '-'
                int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

                if (intNoOfDays <= 44)
                {
                    for (int j = intNoOfDays + 1; j <= 44; j++)
                    {
                        dt.Columns.Add("DAY" + j);
                    }
                }
                else
                {

                }

                #region Change the column names
                //Change the column names to the correct column names.
                Dictionary<string, string> dictNames = new Dictionary<string, string>();
                DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                     "Select * from varnames");
                dictNames.Clear();

                dictNames = TB.loadDict(varNames, dictNames);
                int counter = 0;

                //If it is a column with a date as a name.
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Substring(0, 1) == "2")
                    {
                        if (counter == 0)
                        {
                            strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            if (column.Ordinal == dt.Columns.Count - 1)
                            {

                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;

                            }
                            else
                            {
                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;
                            }
                        }


                    }
                    else
                    {
                        if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                        {
                            column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                        }

                    }
                }

                dt.Columns.Add("FSH");
                dt.Columns.Add("LSH");
                dt.Columns.Add("SECTION");
                dt.Columns.Add("EMPLOYEETYPE");
                dt.Columns.Add("PERIOD");      //xxxxxxxx
                dt.AcceptChanges();


                foreach (DataRow row in dt.Rows)
                {

                    row["FSH"] = strSheetFSH;
                    row["LSH"] = strSheetLSH;
                    row["MININGTYPE"] = "OFFICIALS";
                    row["BONUSTYPE"] = "SERVICES";
                    row["PERIOD"] = BusinessLanguage.Period;   //xxx

                    row["SECTION"] = "OFF";

                    if (row["WAGECODE"].ToString().Trim() == "")
                    {
                        row["WAGECODE"] = "00000";
                    }
                    else
                    {
                    }
                    row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                        {
                            row[i] = "-";
                        }
                    }
                }

                #endregion
                //exportToExcel("c:\\", dt);
                //Write to the database
                TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");

                Application.DoEvents();

                //grdClocked.DataSource = dt;

               // MessageBox.Show("Clockedshifts records are : " + dt.Rows.Count, "INFO", MessageBoxButtons.OK);

                #endregion

            #region UPDATE the participants
                //update the table participants with the correct section.
                //The correct section's FSH and LSH must be saved.
                
                DataTable temp = new DataTable();
                DataTable temp2 = new DataTable();
                
                foreach (DataRow dr in Participants.Rows)
                {
                    IEnumerable<DataRow> query5 = from locks in Clocked.AsEnumerable()
                                                  where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == dr["EMPLOYEE_NO"].ToString().Trim()
                                                  select locks;

                    try
                    {
                        temp = query5.CopyToDataTable<DataRow>();
                        //Change only if section = 'XXX'
                        if (dr["SECTION"].ToString().Trim() == "XXX")
                        {
                            dr["MEASSECTION"] = temp.Rows[0]["MEASSECTION"];
                        }
                        IEnumerable<DataRow> query6 = from locks in Calendar.AsEnumerable()
                                                      where locks.Field<string>("SECTION").TrimEnd() == temp.Rows[0]["MEASSECTION"].ToString().Trim()
                                                      select locks;
                        try
                        {
                            temp2 = query6.CopyToDataTable<DataRow>();
                            //Change only if = 0
                            if (dr["FSH"].ToString().Trim() == "0")
                            {
                                dr["FSH"] = temp2.Rows[0]["FSH"];
                            }
                            if (dr["LSH"].ToString().Trim() == "0")
                            {
                                dr["LSH"] = temp2.Rows[0]["LSH"];
                            }
                            if (dr["MONTHSHIFTS"].ToString().Trim() == "0")
                            {
                                dr["MONTHSHIFTS"] = temp2.Rows[0]["MONTHSHIFTS"];
                            }

                        }
                        catch
                        {
                            IEnumerable<DataRow> query7 = from locks in Calendar.AsEnumerable()
                                                          where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                                          select locks;

                            try
                            {
                                temp = query7.CopyToDataTable<DataRow>();
                                //Change only if = 0
                                if (dr["FSH"].ToString().Trim() == "0")
                                {
                                    dr["FSH"] = temp.Rows[0]["FSH"];
                                }
                                if (dr["LSH"].ToString().Trim() == "0")
                                {
                                    dr["LSH"] = temp.Rows[0]["LSH"];
                                }
                                if (dr["MONTHSHIFTS"].ToString().Trim() == "0")
                                {
                                    dr["MONTHSHIFTS"] = temp.Rows[0]["MONTHSHIFTS"];
                                }
                            }
                            catch
                            {
                                 IEnumerable<DataRow> query8 = from locks in Calendar.AsEnumerable()
                                                          where locks.Field<string>("SECTION").TrimEnd() == "OFF"
                                                          select locks;

                                 try
                                 {
                                     temp = query8.CopyToDataTable<DataRow>();
                                     //Change only if = 0
                                     if (dr["FSH"].ToString().Trim() == "0")
                                     {
                                         dr["FSH"] = temp.Rows[0]["FSH"];
                                     }
                                     if (dr["LSH"].ToString().Trim() == "0")
                                     {
                                         dr["LSH"] = temp.Rows[0]["LSH"];
                                     }
                                     if (dr["MONTHSHIFTS"].ToString().Trim() == "0")
                                     {
                                         dr["MONTHSHIFTS"] = temp.Rows[0]["MONTHSHIFTS"];
                                     }
                                 }
                                 catch
                                 {

                                 }
                            }
                        }
                    }
                    catch
                    {
                        IEnumerable<DataRow> query8 = from locks in Calendar.AsEnumerable()
                                                      where locks.Field<string>("SECTION").TrimEnd() == "OFF"
                                                      select locks;

                        try
                        {
                            temp = query8.CopyToDataTable<DataRow>();
                            //Change only if = 0
                            dr["SECTION"] = "OFF";      
                            dr["FSH"] = temp.Rows[0]["FSH"];                      
                            dr["LSH"] = temp.Rows[0]["LSH"];
                            dr["MONTHSHIFTS"] = temp.Rows[0]["MONTHSHIFTS"];

                        }
                        catch
                        {

                        }


                    }

                    Participants.AcceptChanges();
                }

                Participants.AcceptChanges();
                grdParticipants.DataSource = Participants;
                TB.saveCalculations2(Participants, Base.DBConnectionString, "", "PARTICIPANTS");
                 
                string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                    "substring(gang,1,5) as MEASSECTION from Clockedshifts where employee_no in (select distinct employee_no from participants)";

                BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                //Update bonusshifts with the correct FSH and LSH extracted from participants
                //The FSH and LSH is updatable by the Mining Manager and therefore it is taken from
                //Participants and not from calendar.
                 
                foreach (DataRow dr in BonusShifts.Rows)
                {
                    IEnumerable<DataRow> query10 = from locks in Participants.AsEnumerable()
                                                  where locks.Field<string>("Employee_No").TrimEnd() == dr["EMPLOYEE_NO"].ToString().Trim()
                                                  select locks;
                    try
                    {
                        temp2 = query10.CopyToDataTable<DataRow>();
                        dr["FSH"] = temp2.Rows[0]["FSH"];
                        dr["LSH"] = temp2.Rows[0]["LSH"];
                        dr["MONTHSHIFTS"] = temp2.Rows[0]["MONTHSHIFTS"];
                    }
                    catch
                    {
                    }
                }

                BonusShifts.AcceptChanges();
                 
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts
                //Then count each employee's shifts individually 
                
                foreach (DataRow dr in BonusShifts.Rows)
                {

                    string strCalendarFSH = dr["FSH"].ToString();
                    string strCalendarLSH = dr["LSH"].ToString();
                    if (strCalendarFSH.Substring(0,2) == "20")
                    {
                        DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                        DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                        sheetfhs = SheetFSH;
                        sheetlhs = SheetLSH;
                        int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
                        int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
                        int intStopDay = 0;

                        #region Check FSH and LSH differance
                        if (intStartDay < 0)
                        {
                            //The calendarFSH falls outside the startdate of the sheet.
                            intStartDay = 0;
                        }
                        else
                        {

                        }

                        if (intEndDay < 0 && intEndDay < -44)
                        {
                            intStopDay = 0;
                        }
                        else
                        {
                            if (intEndDay < 0)
                            {
                                //the LSH of the measuring period falls within the spreadsheet
                                intStopDay = intNoOfDays + intEndDay;

                            }
                            else
                            {
                                //The LSH of the measuring period falls outside the spreadsheet
                                intStopDay = 44;
                            }


                            //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                            //were not imported.
                        #endregion

                        #region count the shifts
                            //Count the the shifts


                            extractAndCalcShifts(intStartDay, intStopDay, dr);
                            

                            }

                            #endregion

                       
                
                     }
                   
                    #endregion
                }

                //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                DataColumn dcPeriod = new DataColumn();
                dcPeriod.ColumnName = "PERIOD";
                BonusShifts.Columns.Remove("PERIOD");
                BonusShifts.AcceptChanges();
                InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

                foreach (DataRow dr in BonusShifts.Rows)
                {
                    dr["PERIOD"] = BusinessLanguage.Period;
                }

                TB.saveCalculations2(BonusShifts, Base.DBConnectionString, "", "BONUSSHIFTS");
                
                this.Cursor = Cursors.Arrow;
                File.Delete(FilePath);
                }
            }

        private void UpdateClockedShifts()
        {
            //This process runs only when clocked shifts are imported.

            #region Extract dates
            //Load the section's first and last shift date
            DateTime dteFSH = dateTimePicker1.Value;
            //DateTime dteLSH = dateTimePicker2.Value;

            string tempdte = Clocked.Rows[1]["FSH"].ToString().Trim();
            //DateTime dteDateFrom = Convert.ToDateTime(tempdte.Trim());

            tempdte = Clocked.Rows[1]["LSH"].ToString().Trim();
            //DateTime dteDateEnd = Convert.ToDateTime(tempdte.Trim());

            //int intstart = dteDateFrom.Subtract(dteFSH).Days + 1;
            //int intend = dteLSH.Subtract(dteDateFrom).Days + 2;

            #endregion

            //foreach (DataRow dr in Offdays.Rows)
            //{
            //    string offdate = dr["OFFDAYVALUE"].ToString();
            //    if (offdate.Trim() == "2009-01-01")
            //    {
            //    }
            //    else
            //    {

            //        DateTime dteOffdate = Convert.ToDateTime(dr["OFFDAYVALUE"].ToString());

            //        int intOffday = dteOffdate.Subtract(dteFSH).Days;

            //        Base.UpdateOffdays(Base.DBConnectionString, intOffday);

            //        Application.DoEvents();
            //    }
            //}

        }

        public void InsertAfter(DataColumnCollection columns, DataColumn currentColumn, DataColumn newColumn)
        {
            if (columns.Contains(currentColumn.ColumnName))
            {
                columns.Add(newColumn);
                //add the new column after the current one 
                columns[newColumn.ColumnName].SetOrdinal(currentColumn.Ordinal + 1);
            }
            else
            {
                throw new ArgumentException(/** snip **/);
            }
        }

        private void extractAndCalcShifts(int DayStart, int DayEnd,DataRow BonusRow)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int shiftsCheck = 0;

                foreach (DataColumn column in BonusShifts.Columns)
                {
                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        int i = BonusShifts.Columns[column.ColumnName].Ordinal;
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd) &&
                           DayStart <= DayEnd))
                        {
                            if (BonusRow[i].ToString().Trim() == "U" || BonusRow[i].ToString().Trim() == "u" || BonusRow[i].ToString().Trim() == "T" || 
                                BonusRow[i].ToString().Trim() == "W" || BonusRow[i].ToString().Trim() == "r" ||
                                BonusRow[i].ToString().Trim() == "z")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                                shiftsCheck = 1;    
                            }
                            else
                            {
                                if (BonusRow[i].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else { }

                            }
                        }
                        else
                        {
                            BonusRow[i] = "*";
                        }
                    }
                    else
                    {
                       
                    }
                }//foreach datacolumn

                BonusRow["SHIFTS_WORKED"] = intShiftsWorked;
                BonusRow["AWOP_SHIFTS"] = intAwopShifts;
                intShiftsWorked = 0;
                intAwopShifts = 0;

                BonusShifts.AcceptChanges();

            Application.DoEvents();

        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            lstBErrorLog.Items.Clear();
            //string faultPath = "c:\\Reports\\BonusshiftsvsClockedShiftsReport" + DateTime.Now.ToString("yymmddhhmmss") + ".txt";//Sets the path file
            //StreamWriter sw = new StreamWriter(faultPath);
            int goOn = 1;

            string strProcess = tabInfo.SelectedTab.Name;

            if (strProcess == "tabLabour")
            {
              

                goOn = 1;

 
            }

            if (strProcess == "tabCalendar")
            {
                if (btnLock.Text == "Lock")
                {
                    TB.InsertData(Base.DBConnectionString, "UPDATE Participants SET Participants.FSH = Calendar.FSH,Participants.LSH = Calendar.LSH " +
                                                  " FROM Participants  INNER JOIN  Calendar ON Participants.MeasSection = Calendar.Section " +
                                                  " and Participants.LSH = '0' and Participants.PERIOD = Calendar.PERIOD AND Calendar.Period = '" + BusinessLanguage.Period.Trim() + "'");
                    Application.DoEvents();
                    TB.InsertData(Base.DBConnectionString, "UPDATE Participants SET Participants.FSH = Calendar.FSH,Participants.LSH = Calendar.LSH " +
                                                  " FROM Participants  INNER JOIN  Calendar ON Participants.Section = Calendar.Section " +
                                                  " and Participants.LSH = '0' and Participants.PERIOD = Calendar.PERIOD AND Calendar.Period = '" + BusinessLanguage.Period.Trim() + "'");

                    MessageBox.Show("Please re-import the shifts.", "Information", MessageBoxButtons.OK);
                }
                else
                {
                    TB.InsertData(Base.DBConnectionString, "UPDATE Participants SET Participants.FSH = '0', Participants.LSH = '0' where period = '" + BusinessLanguage.Period.Trim() + "';");
                }
            }


            if (goOn == 1)
            {

                if (btnLock.Text == "Lock")
                {

                    TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                          "' and period = '" + txtPeriod.Text.Trim() + "' and section = 'OFF'");
                    btnLock.Text = "Unlock";

                }

                else
                {

                    TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = '" + strProcess +
                                          "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                    btnLock.Text = "Lock";

                }

                evaluateInputProcessStatus();
                listBox2_SelectedIndexChanged("Method", null);
                openTab(tabProcess);

                Application.DoEvents();
            }
            else
            {

                MessageBox.Show("Bonus shifts cannot be more than Clocked shifts!", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Error);
     
                string resultDialog = MessageBox.Show("Do you want to print the report?", "Print", MessageBoxButtons.YesNo).ToString();
                //This will print the report to the default printer installed
                if (resultDialog.ToString() == "Yes")
                {
                    string printPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    System.Drawing.Font printFont = new System.Drawing.Font("Arial", 10);
                    PrintDocument printDocument1 = new PrintDocument();
 
                    printDocument1.Print();

                }
            }

        }

        private void grdMiners_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            MessageBox.Show("column = " + e.ColumnIndex + "  row = " + e.RowIndex);
            string test = scrOfficials.ActiveForm.ActiveControl.Name;
            MessageBox.Show("name = " + test);
        }

        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            string strSQL = string.Empty;
            string strName = string.Empty;
            string strDesignation = string.Empty;
            string strDesignationDesc = string.Empty;

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabParticipants":
                    #region tabParticipants
                    if (cboNames.Text.Trim().Length > 0 &&
                        cboParticipantsDesignation.Text.Trim().Length > 0 )
                    {

                        if (cboParticipantsDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboParticipantsDesignation.Text.Substring(0, cboParticipantsDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboParticipantsDesignation.Text.Substring((cboParticipantsDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboParticipantsDesignation.Text.Trim();
                            strDesignationDesc = cboParticipantsDesignation.Text.Trim();
                        }

                        DataTable temp = new DataTable();
                        temp = Participants.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        dr["HOD"] = cboParticipantsHOD.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboNames.Text.Trim();
                        dr["DESIGNATION"] = strDesignation;
                        dr["SECTION"] = "OFF";
                        dr["PERIOD"] = BusinessLanguage.Period.Trim();
                        dr["DESIGNATION_DESC"] = strDesignationDesc;
                        dr["EMPLOYEE_NAME"] = cboMinersEmpName.Text.Trim();
                        dr["MEASSECTION"] = txtParticipantsSection.Text.Trim();
                        dr["FSH"] = Convert.ToDateTime(dateTimePicker3.Text).ToString("yyyy-MM-dd");
                        dr["LSH"] = Convert.ToDateTime(dateTimePicker4.Text).ToString("yyyy-MM-dd");
                        dr["MONTHSHIFTS"] = cboParticipantsMeasShifts.Text.Trim();
                        dr["VALIDITY"] = cboValidity.Text.Trim();
                        temp.Rows.Add(dr);
                        string strDelete = " where employee_no = '999' ";
                        int rowindex = grdParticipants.CurrentCell.RowIndex;
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "Participants");
                        evaluateParticipants();
                        Application.DoEvents();

                        for(int i = 0; i <= grdParticipants.Rows.Count - 1; i++)
                        {
                            if(grdParticipants["EMPLOYEE_NO",i].Value.ToString().Trim() == cboNames.Text.Trim())
                            {
                                for (int j = 0; j <= grdParticipants.Columns.Count - 1; j++)
                                {
                                    grdParticipants[j, i].Style.BackColor = Color.Plum;
                                }
                                i = grdParticipants.Rows.Count;
                            }
                        }   
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabOfficials":
                    #region tabOfficials
                    //HJ
                    if (cboOfficialsEmployeeno.Text.Trim().Length != 0 && cboOfficialsDesignation.Text.Trim().Length != 0 &&
                        txtSafetyActual.Text.Trim().Length != 0 && txtCostActual.Text.Trim().Length != 0 &&
                        txtCostPlanned.Text.Trim().Length != 0 && txtGoldActual.Text.Trim().Length != 0 &&
                         txtGoldPlanned.Text.Trim().Length != 0 && txtProductionActual.Text.Trim().Length != 0 &&
                        txtProductionPlanned.Text.Trim().Length != 0)
                    {

                        if (cboOfficialsDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboOfficialsDesignation.Text.Substring(0, cboOfficialsDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboOfficialsDesignation.Text.Substring((cboOfficialsDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboOfficialsDesignation.Text.Trim();
                            strDesignationDesc = cboOfficialsDesignation.Text.Trim();
                        }

                        DataTable temp = new DataTable();
                        temp = Officials.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        //Extract the employee to be inserted record from bonusshifts
                        IEnumerable<DataRow> query1 = from locks in Labour.AsEnumerable()
                                                      where locks.Field<string>("Employee_No").TrimEnd() == cboOfficialsEmployeeno.Text.Trim()
                                                      select locks;


                        DataTable tempEmpl = query1.CopyToDataTable<DataRow>();

                        if (tempEmpl.Rows.Count > 0)
                        {


                            dr["EMPLOYEE_NO"] = cboOfficialsEmployeeno.Text.Trim();
                            dr["EMPLOYEE_NAME"] = txtOfficialsName.Text.Trim();
                            dr["GANG"] = tempEmpl.Rows[0]["GANG"];
                            dr["WAGECODE"] = tempEmpl.Rows[0]["Wagecode"];
                            dr["WAGE_DESCRIPTION"] = tempEmpl.Rows[0]["Wage_Description"];
                            dr["ACTING"] = tempEmpl.Rows[0]["ACTING"];
                            dr["SUD"] = tempEmpl.Rows[0]["SUD"];
                            dr["PROCESSCODE"] = tempEmpl.Rows[0]["PROCESSCODE"];
                            dr["LINERESPCODE"] = tempEmpl.Rows[0]["LINERESPCODE"];
                            dr["MININGTYPE"] = "OFFICIALS";
                            dr["BONUSTYPE"] = "SERVICES";
                            dr["DAY0"] = "X";
                            dr["DAY1"] = "X";
                            dr["DAY2"] = "X";
                            dr["DAY3"] = "X";
                            dr["DAY4"] = "X";
                            dr["DAY5"] = "X";
                            dr["DAY6"] = "X";
                            dr["DAY7"] = "X";
                            dr["DAY8"] = "X";
                            dr["DAY9"] = "X";
                            dr["DAY10"] = "X";
                            dr["DAY11"] = "X";
                            dr["DAY12"] = "X";
                            dr["DAY13"] = "X";
                            dr["DAY14"] = "X";
                            dr["DAY15"] = "X";
                            dr["DAY16"] = "X";
                            dr["DAY17"] = "X";
                            dr["DAY18"] = "X";
                            dr["DAY19"] = "X";
                            dr["DAY20"] = "X";
                            dr["DAY21"] = "X";
                            dr["DAY22"] = "X";
                            dr["DAY23"] = "X";
                            dr["DAY24"] = "X";
                            dr["DAY25"] = "X";
                            dr["DAY26"] = "X";
                            dr["DAY27"] = "X";
                            dr["DAY28"] = "X";
                            dr["DAY29"] = "X";
                            dr["DAY30"] = "X";
                            dr["DAY31"] = "X";
                            dr["DAY32"] = "X";
                            dr["DAY33"] = "X";
                            dr["DAY34"] = "X";
                            dr["DAY35"] = "X";
                            dr["DAY36"] = "X";
                            dr["DAY37"] = "X";
                            dr["DAY38"] = "X";
                            dr["DAY39"] = "X";
                            dr["DAY40"] = "X";
                            dr["DAY41"] = "X";
                            dr["DAY42"] = "X";
                            dr["DAY43"] = "X";
                            dr["DAY44"] = "X";
                            dr["FSH"] = tempEmpl.Rows[0]["FSH"];
                            dr["LSH"] = tempEmpl.Rows[0]["LSH"];
                            dr["SECTION"] = tempEmpl.Rows[0]["SECTION"];
                            dr["EMPLOYEETYPE"] = tempEmpl.Rows[0]["EMPLOYEETYPE"];
                            dr["SHIFTS_WORKED"] = tempEmpl.Rows[0]["SHIFTS_WORKED"];
                            dr["AWOP_SHIFTS"] = tempEmpl.Rows[0]["AWOP_SHIFTS"];
                            dr["STRIKE_SHIFTS"] = tempEmpl.Rows[0]["STRIKE_SHIFTS"];
                            dr["DESIGNATION"] = strDesignation.TrimEnd();
                            dr["DESIGNATION_DESC"] = strDesignationDesc.TrimStart().TrimEnd();
                            dr["SAFETY_ACTUAL"] = txtSafetyActual.Text.Trim();
                            dr["COST_ACTUAL"] = txtCostActual.Text.Trim();
                            dr["COST_PLANNED"] = txtCostPlanned.Text.Trim();
                            dr["GOLD_ACTUAL"] = txtGoldActual.Text.Trim();
                            dr["GOLD_PLANNED"] = txtGoldPlanned.Text.Trim();
                            dr["PRODUCTION_ACTUAL"] = txtProductionPlanned.Text.Trim();
                            dr["PRODUCTION_PLANNED"] = txtProductionPlanned.Text.Trim();

                            temp.Rows.Add(dr);
                            //Create a total invalid delete.
                            string strDelete = " where Employee_no = '999'";
                            int rowindex = grdOfficials.CurrentCell.RowIndex;
                            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "OFFICIALS");
                            evaluateOfficials();

                            grdOfficials.FirstDisplayedScrollingRowIndex = rowindex;
                        }
                        else
                        {

                            MessageBox.Show("Insert is impossible, because employee does not exist on BonusShifts", "Information", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties
                    if (cboEmplPenEmployeeNo.Text.Trim().Length > 0 &&
                        txtPenaltyValue.Text.Trim().Length > 0 && cboPenaltyInd.Text.Trim().Length > 0)
                    {
                        DataRow dr;
                        dr = EmplPen.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboEmplPenEmployeeNo.Text.Trim();
                        dr["PENALTYVALUE"] = txtPenaltyValue.Text.Trim();
                        dr["PENALTYIND"] = cboPenaltyInd.Text.Trim();

                        EmplPen.Rows.Add(dr);

                        strSQL = "Insert into EmployeePenalties values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + cboEmplPenEmployeeNo.Text.Trim() + "', '" + txtPenaltyValue.Text.Trim() +
                                 "', '" + cboPenaltyInd.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion
              
                case "tabRates":
                    #region tabRates
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {
                        DataRow dr;
                        dr = Rates.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["RATE_TYPE"] = txtRateType.Text.Trim();
                        dr["LOW_VALUE"] = txtLowValue.Text.Trim();
                        dr["HIGH_VALUE"] = txtHighValue.Text.Trim();
                        dr["RATE"] = txtRate.Text.Trim();

                        int rowindex = grdRates.CurrentCell.RowIndex;
                        strSQL = "Insert into Rates values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtRateType.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + txtLowValue.Text.Trim() + "', '" + txtHighValue.Text.Trim() +
                                 "', '" + txtRate.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        
                        grdRates.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

            }
        }

        private string checkSQL(int intCounter, string strSQL)
        {
            if (intCounter > 0)
            {
                for (int i = 0; i <= intCounter - 1; i++)
                {
                    strSQL = strSQL.Trim() + ",'0'";
                }
                strSQL = strSQL.Trim() + ")";
            }
            else
            {
                strSQL = strSQL.Trim() + "')";
            }

            return strSQL;
        }

        
        private void tabInfo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text == "***")
            {
                MessageBox.Show("Please select a section.", "Information", MessageBoxButtons.OK);
            }
            else
            {
                btnInsertRow.Enabled = true;
                btnUpdate.Enabled = true;

                btnDeleteRow.Enabled = false;
                listBox1.Enabled = false;                               //HJ
                btnLoad.Enabled = false;
                dateTimePicker1.Enabled = false;                        //HJ
                dateTimePicker2.Enabled = false;                        //HJ
                btnPrint.Enabled = false;
                btnLock.Enabled = false;
                panelLock.BackColor = Color.Lavender;

                int intCount = checkLock(tabInfo.SelectedTab.Name);
                if (intCount > 0)
                {
                    btnLock.Text = "Unlock";
                }
                else
                {
                    btnLock.Text = "Lock";
                }

                switch (tabInfo.SelectedTab.Name)
                {
                    #region tabCalendar
                    case "tabCalendar":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnLoad.Enabled = true;
                        dateTimePicker1.Enabled = true;                 
                        dateTimePicker2.Enabled = true;                 
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        label36.Visible = false;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;


                        break;
                    #endregion

                    #region tabClockShifts
                    case "tabClockShifts":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnPrint.Enabled = true;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;
                    #endregion

                    #region tabLabour
                    case "tabLabour":

                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateLabour();
                        break;
                    #endregion

                    #region tabOfficials
                    case "tabOfficials":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender; 
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateOfficials();
                        break;
                    #endregion

                    #region tabParticipants
                    case "tabParticipants":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateParticipants();
                        break;
                    #endregion

                    #region tabConfig
                    case "tabConfig":

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabEmplPen
                    case "tabEmplPen":

                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        break;

                    #endregion

                    #region tabSelected
                    case "tabSelected":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        listBox1.Enabled = true;                            //HJ
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabStatus

                    case "tabProcess":

                        evaluateStatus();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabRates
                    case "tabRates":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        break;

                    #endregion

                }
            }

        }

        private int checkLock(string processToBeChecked)
        {
            //Lynx....LINQ
            DataTable contactTable = TB.getDataTable(TB.TBName);

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "Y"
                                          where locks.Field<string>("PROCESS").TrimEnd() == processToBeChecked
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;
        }

        private int checkLockInputProcesses()
        {
            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("Period").TrimEnd() == BusinessLanguage.Period
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;

        }

        private int checkLockCalendarProcesses()
        {
            //This method returns a 1 if the the current status of calendar is: UNLOCKED
            // or a 0 if the calendar is LOCKED
            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("PROCESS").TrimEnd() == "tabCalendar"
                                          where locks.Field<string>("Period").TrimEnd() == BusinessLanguage.Period
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;
        }

        private void grdEmplPen_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboEmplPenEmployeeNo.Text = grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtPenaltyValue.Text = grdEmplPen["PENALTYVALUE", e.RowIndex].Value.ToString().Trim();
                cboPenaltyInd.Text = grdEmplPen["PENALTYIND", e.RowIndex].Value.ToString().Trim();
                if (grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXXXXXXXXXXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }
            }
            Cursor.Current = Cursors.Arrow;

        }

        private void grdLabour_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboBonusShiftsGang.Text = grdLabour["GANG", e.RowIndex].Value.ToString().Trim(); 
                txtEmployeeNo.Text = grdLabour["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtEmployeeName.Text = grdLabour["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsWageCode.Text = grdLabour["WAGECODE", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsResponseCode.Text = grdLabour["LINERESPCODE", e.RowIndex].Value.ToString().Trim();
                txtShifts.Text = grdLabour["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                txtAwop.Text = grdLabour["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                txtStrikeShifts.Text = grdLabour["STRIKE_SHIFTS", e.RowIndex].Value.ToString().Trim();
                
                
            }

            Cursor.Current = Cursors.Arrow;

        }

        private void grdConfigs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboParameterName.Text = grdConfigs["PARAMETERNAME", e.RowIndex].Value.ToString().Trim();
                cboParm1.Text = grdConfigs["PARM1", e.RowIndex].Value.ToString().Trim();
                cboParm2.Text = grdConfigs["PARM2", e.RowIndex].Value.ToString().Trim();
                cboParm3.Text = grdConfigs["PARM3", e.RowIndex].Value.ToString().Trim();
                cboParm4.Text = grdConfigs["PARM4", e.RowIndex].Value.ToString().Trim();
                cboParm5.Text = grdConfigs["PARM5", e.RowIndex].Value.ToString().Trim();
                cboParm6.Text = grdConfigs["PARM6", e.RowIndex].Value.ToString().Trim();
                cboParm7.Text = grdConfigs["PARM7", e.RowIndex].Value.ToString().Trim();
            }
            Cursor.Current = Cursors.Arrow;
        }

        #region AutoSize
        private void autoSizeGrid(DataGridView DG)
        {
            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader.ToString())
            {
                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                {
                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                }
                else
                {
                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.ColumnHeader.ToString())
                    {
                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    }
                    else
                    {
                        if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                        {
                            DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                        }
                        else
                        {
                            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                            {
                                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                            }
                            else
                            {
                                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.Fill.ToString())
                                {
                                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                                }
                                else
                                {
                                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                                    {
                                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void grdActiveSheet_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdActiveSheet);
            }
        }


        private void grdCalendar_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdCalendar);
            }
        }

        

       

        private void grdLabour_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdLabour);
            }
        }

      

       

        

        private void grdConfigs_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdConfigs);
            }
        }

        private void grdOfficials_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdOfficials);
            }
        }

        private void DoDataExtract()
        {
            connectToDB();
            TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);

        }

        private void DoDataExtract(string Where)
        {
            connectToDB();
            if (Where.Trim().Length == 0)
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);
            }
            else
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName, Where);

            }
        }

        #endregion

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //xxxxxxxxxxxxxxxxxxxxxx
            if (blTablenames == true)
            {

                //extract the data for the selected table.
                string FormulaTableName = string.Empty;
                TB.TBName = (string)listBox1.SelectedItem;
                TB.DBName = Base.DBName;

                if (TB.TBName.Trim().ToUpper().Contains("EARN") && TB.TBName.Trim().ToUpper().Contains("20"))
                {
                    FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));   //xxxxxxxxxxxxxxxxxx
                }
                else
                {
                    FormulaTableName = TB.TBName;
                }

                TB.DBName = Base.DBName;

                connectToDB();
                cboColumnValues.Items.Clear();
                cboColumnNames.Items.Clear();
                cboColumnNames.Text = string.Empty;
                cboColumnValues.Text = string.Empty;

                //Extract distinct column names of the table selected from the list
                _lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, TB.TBName);

                foreach (string s in _lstColumnNames)
                {
                    cboColumnNames.Items.Add(s.Trim());
                }

                TB.ListOfSelectedTableColumns = _lstColumnNames;

                DoDataExtract(strWhere);

                _newDataTable = TB.getDataTable(TB.TBName);
                if (_newDataTable == null)
                {
                    DoDataExtract("");
                    _newDataTable = TB.getDataTable(TB.TBName);

                }


                grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);

                AConn = Analysis.AnalysisConnection;
                AConn.Open();

                _formulas = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(),
                                                                       FormulaTableName, Base.AnalysisConnectionString);

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }

                autoSizeGrid(grdActiveSheet);
            }
            else
            {

            }
        }

        private void exportToExcel(string path, DataTable dt)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = path + "\\" + TB.TBName + ".xls";
                try
                {
                    StreamWriter SW = new StreamWriter(OPath);
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {
                            grid.RenderControl(HTMLWriter);
                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();
                    MessageBox.Show("Your spreadsheet was created at: " + OPath, "Information", MessageBoxButtons.OK);
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void TBExport_Click(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void saveTheSpreadSheet()
        {
            string path = @"c:\" + TB.DBName + "\\" + TB.TBName;
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                DoDataExtract();
                DataTable outputTable = TB.getDataTable(TB.TBName);
                exportToExcel(path, outputTable);
                MessageBox.Show("Successfully Downloaded.", "Information", MessageBoxButtons.OK);

            }
            catch (Exception ee)
            {
                Console.WriteLine("The process failed: {0}", ee.ToString());
            }

            finally { }
        }

        private void grdActiveSheet_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            object intCount = Analysis.countcalcbyname(TB.DBName, TB.TBName, TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);
            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetails(TB.DBName, TB.TBName, TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS");
                DataTable dtFactors = TB.getDataTable("FACTORS");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        MessageBox.Show("SQL extract", "Not available to be tested", MessageBoxButtons.OK);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {
                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void loadDict(DataTable _datatable)
        {
            foreach (DataRow _row in _datatable.Rows)
            {
                string str = _row[0].ToString().Trim();
                if (dict.ContainsKey(str))
                {
                    dict.Remove(str);
                    dict.Add(str, _row[1].ToString().Trim());
                }
                else
                {
                    dict.Add(str, _row[1].ToString().Trim());
                }
            }
            dict.Remove("X");
            dict.Add("X", "0");

        }

        private void buildDisplaySQL(string strwhere, decimal decValue)
        {
            string strSQL = "";

            strSQL = "Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                         TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + TBFormulas.A + TBFormulas.B + TBFormulas.C + TBFormulas.D + TBFormulas.E + TBFormulas.F + TBFormulas.G + TBFormulas.H + " " + strwhere;
            strSQL = strSQL.Replace("#", "").Replace(":and:", "and").Replace(" from ", "\n from ").Replace(" and ", "\n and ").Replace(" where ", "\n where ");

            General.textTestSQL = strSQL;
            scrQuerySQL testsql = new scrQuerySQL();
            testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
            testsql.ShowDialog();

        }

        private void userProfile_Click(object sender, EventArgs e)
        {
            scrProfile userProfile = new scrProfile();
            userProfile.FormLoad(BusinessLanguage, BaseConn);
            userProfile.Show();
        }

        private void grantAccessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scrSecurity useraccess = new scrSecurity();
            useraccess.userAccessLoad(myConn, Base, TB, BusinessLanguage.Userid,strServerPath.ToString().ToUpper());
            useraccess.Show();
        }

       
        private void btnReset_Click(object sender, EventArgs e)
        {
           // grdClocked.DataSource = Clocked;
        }

        private void grdActiveSheet_CellContentDouble_Click(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            string FormulaTableName = string.Empty;

            if (TB.TBName.Trim().ToUpper().Contains("EARN"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            object intCount = Analysis.countcalcbyname(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                       TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);

            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetailsDCript(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                                    TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                //TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS"," Where period = '" + BusinessLanguage.Period + "'");
                DataTable dtFactors = TB.createDataTableWithAdapter(Base.DBConnectionString,
                                    "Select Varname,Varvalue from FACTORS where period = '" + BusinessLanguage.Period + "'");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        string strWhere = " ";
                        for (int i = 0; i < grdActiveSheet.Columns.Count - 1; i++)
                        {

                            strWhere = strWhere.Trim() + " and t1." + grdActiveSheet.Columns[i].HeaderText.Trim() +
                                       " = '" + (string)(grdActiveSheet[i, e.RowIndex].Value).ToString().Trim() + "'";

                        }

                        buildDisplaySQL(strWhere, decValue);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {

                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + BusinessLanguage.Period.Trim() + '\n' + "Table Name:           " + FormulaTableName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        
       

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //xxxxxxxxxxxxxxxxxxxx
            this.Cursor = Cursors.WaitCursor;

            if (listBox2.SelectedIndex >= 0)
            {

                this.Cursor = Cursors.WaitCursor;
                txtSelectedSection.Text = listBox2.SelectedItem.ToString().Trim();
                Base.Section = txtSelectedSection.Text.Trim(); ;   //xxxxxxxxxxxxxxxxxx

                int intRowPosition = 0;
                for (int i = 0; i <= Calendar.Rows.Count - 1; i++)
                {
                    if (Calendar.Rows[i]["SECTION"].ToString() == txtSelectedSection.Text.Trim() &&
                        Calendar.Rows[i]["PERIOD"].ToString() == BusinessLanguage.Period.Trim())
                    {
                        intRowPosition = i;
                    }
                }

                loadDatePickers(intRowPosition);

                cboOffDaysSection.Text = txtSelectedSection.Text.Trim();
                cboOffDaysGang.Text = @"DUMMY";
                label15.Text = listBox2.SelectedItem.ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + listBox2.SelectedItem.ToString().Trim() + "' and period = '" + BusinessLanguage.Period + "'";//xxxxxxxxxxxxxxxxxx

                loadMO();
                extractMeasuringDates();

                evaluateStatus();

                evaluateLabour();
                evaluatePayroll();
                evaluateOfficials();
                evaluateEmployeePenalties();
                this.Cursor = Cursors.Arrow;

                extractMeasuringDates();
            }

        }

        private void extractMeasuringDates()
        {

            IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                          where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();
            dateTimePicker1.Value = Convert.ToDateTime(temp.Rows[0]["FSH"].ToString().Trim());
            dateTimePicker2.Value = Convert.ToDateTime(temp.Rows[0]["LSH"].ToString().Trim());

        }

        private void btnEmployeeCalc_Click(object sender, EventArgs e)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        private void dataSort_Click(object sender, EventArgs e)
        {

        }

        private void DataPrintCrewPrint_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {          
                case "tabLabour":
                    #region tabLabour

                    //HJ
                    if (txtEmployeeNo.Text.Trim().Length > 0
                        && txtEmployeeName.Text.Trim().Length > 0
                        && cboBonusShiftsGang.Text.Trim().Length > 0
                        && cboBonusShiftsWageCode.Text.Trim().Length > 0
                        && cboBonusShiftsResponseCode.Text.Trim().Length > 0
                        && txtShifts.Text.Trim().Length > 0
                        && txtAwop.Text.Trim().Length > 0
                        && txtStrikeShifts.Text.Trim().Length > 0)
                    {

                        intRow = grdLabour.CurrentCell.RowIndex;

                        string strWagecode = Convert.ToString(grdLabour["WAGECODE", intRow].Value);
                        string strEmployeeName = Convert.ToString(grdLabour["EMPLOYEE_NAME", intRow].Value);
                        string strGang = Convert.ToString(grdLabour["GANG", intRow].Value);
                        string strResponseCo = Convert.ToString(grdLabour["LINERESPCODE", intRow].Value);
                        string strShiftsWorked = Convert.ToString(grdLabour["SHIFTS_WORKED", intRow].Value);
                        string strAwops = Convert.ToString(grdLabour["AWOP_SHIFTS", intRow].Value);
                        string strStrikes = Convert.ToString(grdLabour["STRIKE_SHIFTS", intRow].Value);
                        string strDrillerInd = Convert.ToString(grdLabour["DRILLERIND", intRow].Value);
                        string strDrillerShifts = Convert.ToString(grdLabour["DRILLERSHIFTS", intRow].Value);
                        string strTeamLeadind = Convert.ToString(grdLabour["TEAMLEADERIND", intRow].Value);

                        strSQL = "Update bonusshifts set wagecode = '" + cboBonusShiftsWageCode.Text.Trim() +
                                 "' , Gang = '" + cboBonusShiftsGang.Text.Trim() +
                                 "' , Linerespcode = '" + cboBonusShiftsResponseCode.Text.Trim() +
                                 "' , Shifts_Worked = '" + txtShifts.Text.Trim() +
                                 "' , Awop_Shifts = '" + txtAwop.Text.Trim() +
                                 "' , Strike_Shifts = '" + txtStrikeShifts.Text.Trim() +
                                 "' where employee_no = '" + grdLabour["Employee_No", intRow].Value +
                                 "' and Linerespcode = '" + grdLabour["Linerespcode", intRow].Value +
                                 "' and Employee_name = '" + grdLabour["Employee_Name", intRow].Value +
                                 "' and WageCode = '" + grdLabour["WageCode", intRow].Value +
                                 "' and Gang = '" + grdLabour["Gang", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdLabour["WAGECODE", intRow].Value = cboBonusShiftsWageCode.Text.Trim();
                        grdLabour["GANG", intRow].Value = cboBonusShiftsGang.Text.Trim();
                        grdLabour["LINERESPCODE", intRow].Value = cboBonusShiftsResponseCode.Text.Trim();
                        grdLabour["SHIFTS_WORKED", intRow].Value = txtShifts.Text.Trim();
                        grdLabour["AWOP_SHIFTS", intRow].Value = txtAwop.Text.Trim();
                        grdLabour["STRIKE_SHIFTS", intRow].Value = txtStrikeShifts.Text.Trim();

                        for (int i = 0; i <= grdLabour.Columns.Count - 1; i++)
                        {
                            grdLabour[i, intRow].Style.BackColor = Color.Lavender;
                        }
                        //    }
                        //    else
                        //    {
                        //        MessageBox.Show("Invalid password.", "Error", MessageBoxButtons.OK);
                        //    }
                        //}

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion
                   

                case "tabEmplList":

                case "tabOfficials":
                    #region tabOfficials

                    //HJ
                    if (cboOfficialsEmployeeno.Text.Trim().Length != 0 && cboOfficialsDesignation.Text.Trim().Length != 0 &&
                        txtSafetyActual.Text.Trim().Length != 0 && txtCostActual.Text.Trim().Length != 0 &&
                        txtCostPlanned.Text.Trim().Length != 0 && txtGoldActual.Text.Trim().Length != 0 &&
                         txtGoldPlanned.Text.Trim().Length != 0 && txtProductionActual.Text.Trim().Length != 0 &&
                        txtProductionPlanned.Text.Trim().Length != 0)
                    {
                        intRow = grdOfficials.CurrentCell.RowIndex;
                        intColumn = grdOfficials.CurrentCell.ColumnIndex;

                        string strDesignation = string.Empty;
                        string strDesignationDesc = string.Empty;

                        if (cboOfficialsDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboOfficialsDesignation.Text.Substring(0, cboOfficialsDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboOfficialsDesignation.Text.Substring((cboOfficialsDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboOfficialsDesignation.Text.Trim();
                            strDesignationDesc = cboOfficialsDesignation.Text.Trim();
                        }

                        strSQL = "BEGIN transaction; Update Officials set Employee_No = '" + cboOfficialsEmployeeno.Text.Trim() +
                                 "', Designation = '" + strDesignation +
                                 "', Awop_Shifts = '" + txtOfficialsAwops.Text.Trim() +
                                 "', Safety_Actual = '" + txtSafetyActual.Text.Trim() +
                                 "', Gold_Actual = '" + txtGoldActual.Text.Trim() +
                                 "', Gold_Planned = '" + txtGoldPlanned.Text.Trim() +
                                 "', Production_Actual = '" + txtProductionActual.Text.Trim() +
                                 "', Production_Planned = '" + txtProductionPlanned.Text.Trim() +
                                 "', Cost_Actual = '" + txtCostActual.Text.Trim() +
                                 "', Cost_Planned = '" + txtCostPlanned.Text.Trim() +
                                 "', OfflinePercentage = '" + txtOfflinePerc.Text.Trim() +
                                 "', Employee_name = '" + txtOfficialsName.Text.Trim() +
                                 "', Shifts_Worked = '" + txtOfficialsPayshifts.Text.Trim() +
                                 "', Designation_desc = '" + strDesignationDesc.ToString().TrimEnd().TrimStart() + "'" +
                                 " Where Section = '" + grdOfficials["SECTION", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdOfficials["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                 "' and PERIOD = '" + grdOfficials["PERIOD", intRow].Value.ToString().Trim() +
                                 "'; Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        for (int i = 0; i <= grdOfficials.Columns.Count - 1; i++)
                        {
                            grdOfficials[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdOfficials["EMPLOYEE_NO", intRow].Value = cboOfficialsEmployeeno.Text.Trim();
                        grdOfficials["DESIGNATION", intRow].Value = strDesignation;
                        grdOfficials["DESIGNATION_DESC", intRow].Value = strDesignationDesc;
                        grdOfficials["EMPLOYEE_NAME", intRow].Value = txtOfficialsName.Text.Trim();
                        grdOfficials["SAFETY_ACTUAL", intRow].Value = txtSafetyActual.Text.Trim();
                        grdOfficials["COST_ACTUAL", intRow].Value = txtCostActual.Text.Trim();
                        grdOfficials["COST_PLANNED", intRow].Value = txtCostPlanned.Text.Trim();
                        grdOfficials["GOLD_ACTUAL", intRow].Value = txtGoldActual.Text.Trim();
                        grdOfficials["GOLD_PLANNED", intRow].Value = txtGoldPlanned.Text.Trim();
                        grdOfficials["PRODUCTION_ACTUAL", intRow].Value = txtProductionActual.Text.Trim();
                        grdOfficials["PRODUCTION_PLANNED", intRow].Value = txtProductionPlanned.Text.Trim();
                        grdOfficials["OFFLINEPERCENTAGE", intRow].Value = txtOfflinePerc.Text.Trim();
                        grdOfficials["SHIFTS_WORKED", intRow].Value = txtOfficialsPayshifts.Text.Trim();
                        grdOfficials["AWOP_SHIFTS", intRow].Value = txtOfficialsAwops.Text.Trim();

                        grdOfficials.FirstDisplayedScrollingRowIndex = intRow;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabParticipants":
                    #region tabParticipants

                    //HJ
                    if (cboNames.Text.Trim().Length != 0 && cboMinersEmpName.Text.Trim().Length != 0 &&
                        cboParticipantsDesignation.Text.Trim().Length != 0 &&
                        txtParticipantsSection.Text.Trim().Length != 0 && cboValidity.Text.Trim().Length != 0)
                    {
                        intRow = grdParticipants.CurrentCell.RowIndex;
                        intColumn = grdParticipants.CurrentCell.ColumnIndex;

                        string strDesignation = string.Empty;
                        string strDesignationDesc = string.Empty;

                        if (cboParticipantsDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboParticipantsDesignation.Text.Substring(0, cboParticipantsDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboParticipantsDesignation.Text.Substring((cboParticipantsDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboParticipantsDesignation.Text.Trim();
                            strDesignationDesc = cboParticipantsDesignation.Text.Trim();
                        }
                     
                        strSQL = "BEGIN transaction; Update Participants set Designation = '" + strDesignation.Trim() +
                                 "', DESIGNATION_DESC = '" + strDesignationDesc.Trim() +
                                 "', HOD = '" + cboParticipantsHOD.Text.Trim() + 
                                 "', SECTION = '" + txtSelectedSection.Text.Trim() +
                                 "', MEASSECTION = '" + txtParticipantsSection.Text.Trim() + 
                                 "', FSH = '" + Convert.ToDateTime(dateTimePicker3.Text).ToString("yyyy-MM-dd") +
                                 "', LSH = '" + Convert.ToDateTime(dateTimePicker4.Text).ToString("yyyy-MM-dd") +
                                 "', MONTHSHIFTS = '" + cboParticipantsMeasShifts.Text.Trim() +
                                 "', VALIDITY = '" + cboValidity.Text.Trim() +
                                 "', Employee_name = '" + cboMinersEmpName.Text.Trim() +
                                 "' Where Section = '" + grdParticipants["SECTION", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdParticipants["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                 "' and validity = '" + grdParticipants["VALIDITY", intRow].Value.ToString() +
                                 "' and HOD = '" + grdParticipants["HOD", intRow].Value.ToString() +
                                 "' and DESIGNATION_DESC = '" + grdParticipants["DESIGNATION_DESC", intRow].Value.ToString() +
                                 "' and PERIOD = '" + grdParticipants["PERIOD", intRow].Value.ToString().Trim() +
                                 "';Commit Transaction;";


                        TB.InsertData(Base.DBConnectionString, strSQL);

                        for (int i = 0; i <= grdParticipants.Columns.Count - 1; i++)
                        {
                            grdParticipants[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdParticipants["DESIGNATION", intRow].Value = strDesignation;
                        grdParticipants["HOD", intRow].Value = cboParticipantsHOD.Text.Trim();
                        grdParticipants["DESIGNATION_DESC", intRow].Value = strDesignationDesc;
                        grdParticipants["MONTHSHIFTS", intRow].Value = cboParticipantsMeasShifts.Text.Trim();
                        grdParticipants["MEASSECTION", intRow].Value = txtParticipantsSection.Text.Trim();
                        grdParticipants["EMPLOYEE_NAME", intRow].Value = cboMinersEmpName.Text.Trim();
                        grdParticipants["VALIDITY", intRow].Value = cboValidity.Text.Trim();
                        grdParticipants["FSH", intRow].Value = Convert.ToDateTime(dateTimePicker3.Text).ToString("yyyy-MM-dd");
                        grdParticipants["LSH", intRow].Value = Convert.ToDateTime(dateTimePicker4.Text).ToString("yyyy-MM-dd");

                        grdParticipants.FirstDisplayedScrollingRowIndex = intRow;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    //HJ
                    if (cboEmplPenEmployeeNo.Text.Trim().Length != 0 &&
                        txtPenaltyValue.Text.Trim().Length != 0 && cboPenaltyInd.Text.Trim().Length != 0)
                    {

                        intRow = grdEmplPen.CurrentCell.RowIndex;
                        intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                        if (cboEmplPenEmployeeNo.Text.Contains("-"))
                        {
                            strName = cboEmplPenEmployeeNo.Text.Substring(0, cboEmplPenEmployeeNo.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboEmplPenEmployeeNo.Text.Trim();
                        }

                        strSQL = "BEGIN transaction; Update EmployeePenalties set Period = '" + txtPeriod.Text.Trim() +
                                             "', Employee_No = '" + strName + "', PenaltyValue = '" + txtPenaltyValue.Text.Trim() +
                                             "', PenaltyInd = '" + cboPenaltyInd.Text.Trim() + "'" +
                                             " Where Section = '" + grdEmplPen["SECTION", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdEmplPen["PERIOD", intRow].Value.ToString().Trim() +
                                             "' and Employee_No = '" + grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                             "' and PenaltyValue = '" + grdEmplPen["PENALTYVALUE", intRow].Value.ToString().Trim() +
                                             "' and PenaltyInd = '" + grdEmplPen["PENALTYIND", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXXXXXXXXXXX")
                        {
                            grdEmplPen["Section", intRow].Value = txtSelectedSection.Text.Trim();
                            grdEmplPen["Section", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Period", intRow].Value = txtPeriod.Text.Trim();
                            grdEmplPen["Period", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Employee_No", intRow].Value = cboEmplPenEmployeeNo.Text.Trim();
                            grdEmplPen["Employee_No", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyValue", intRow].Value = txtPenaltyValue.Text.Trim();
                            grdEmplPen["PenaltyValue", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyInd", intRow].Value = cboPenaltyInd.Text.Trim();
                            grdEmplPen["PenaltyInd", intRow].Style.BackColor = Color.Lavender;

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            clearAllCalcValues("Officials", txtSelectedSection.Text.Trim()); 

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion


                case "tabConfig":
                    #region tabConfiguration

                    //HJ
                    if (grdConfigs[0, intRow].Value.ToString().Trim() != "XXX")
                    {
                        if (cboParameterName.Text.Trim().Length != 0 && cboParm1.Text.Trim().Length != 0 &&
                            cboParm2.Text.Trim().Length != 0 && cboParm3.Text.Trim().Length != 0 &&
                            cboParm4.Text.Trim().Length != 0 && cboParm5.Text.Trim().Length != 0 &&
                            cboParm6.Text.Trim().Length != 0 && cboParm7.Text.Trim().Length != 0)
                        {

                            intRow = grdConfigs.CurrentCell.RowIndex;
                            intColumn = grdConfigs.CurrentCell.ColumnIndex;

                            InputBoxResult intresult = InputBox.Show("Password: ");

                            if (intresult.ReturnCode == DialogResult.OK)
                            {
                                if (intresult.Text.Trim() == "Moses")
                                {

                                    General.updateConfigsRecord(Base.BaseConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                                     cboParameterName.Text.Trim(), cboParm1.Text.Trim(), cboParm2.Text.Trim(), cboParm3.Text.Trim(), cboParm4.Text.Trim(),
                                     cboParm5.Text.Trim(), cboParm6.Text.Trim(), cboParm7.Text.Trim(), grdConfigs["ParameterName", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm1", intRow].Value.ToString().Trim(), grdConfigs["Parm2", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm3", intRow].Value.ToString().Trim(), grdConfigs["Parm4", intRow].Value.ToString().Trim());
                                    //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                    foreach (string s in lstTableColumns)
                                    {
                                        if (dictGridValues[s] == grdConfigs[s, intRow].Value.ToString().Trim())
                                        {

                                        }
                                        else
                                        {
                                            //Write out to audit log
                                            writeAudit("CONFIGURATION", "U - Update", s, dictGridValues[s], grdConfigs[s, intRow].Value.ToString().Trim());

                                        }

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Invalid password", "Error", MessageBoxButtons.OK);
                                }
                            }

                            grdConfigs["ParameterName", intRow].Value = cboParameterName.Text.Trim();
                            grdConfigs["ParameterName", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm1", intRow].Value = cboParm1.Text.Trim();
                            grdConfigs["Parm1", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm2", intRow].Value = cboParm2.Text.Trim();
                            grdConfigs["Parm2", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm3", intRow].Value = cboParm3.Text.Trim();
                            grdConfigs["Parm3", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm4", intRow].Value = cboParm4.Text.Trim();
                            grdConfigs["Parm4", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm5", intRow].Value = cboParm5.Text.Trim();
                            grdConfigs["Parm5", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm6", intRow].Value = cboParm6.Text.Trim();
                            grdConfigs["Parm6", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm7", intRow].Value = cboParm7.Text.Trim();
                            grdConfigs["Parm7", intRow].Style.BackColor = Color.Lavender;

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion



                case "tabRates":
                    #region tabRates

                    //HJ
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {

                        InputBoxResult result = InputBox.Show("Password: ", "Rates Inputs are Password Protected!", "*", "0");

                        if (result.ReturnCode == DialogResult.OK)
                        {
                            if (result.Text.Trim() == "Moses")
                            {
                                intRow = grdRates.CurrentCell.RowIndex;
                                intColumn = grdRates.CurrentCell.ColumnIndex;

                                General.updateRatesRecord(Base.DBConnectionString, BusinessLanguage.BussUnit,
                                                             txtMiningType.Text.Trim(),
                                                             txtBonusType.Text.Trim(),
                                                             txtPeriod.Text.ToString().Trim(),
                                                             txtRateType.Text.Trim(),
                                                             txtLowValue.Text.Trim(),
                                                             txtHighValue.Text.Trim(), txtRate.Text.Trim(),
                                                             grdRates["Low_Value", intRow].Value.ToString().Trim(),
                                                             grdRates["High_Value", intRow].Value.ToString().Trim(),
                                                             grdRates["Rate", intRow].Value.ToString().Trim());
                                Application.DoEvents();

                                MessageBox.Show("All calculations will becleared.  Recalculations have to be done.", "Information", MessageBoxButtons.OK);
                                clearAllCalcValues("Ganglink", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Miners", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Bonusshifts", txtSelectedSection.Text.Trim());

                                grdRates["Low_Value", intRow].Value = txtLowValue.Text.Trim();
                                grdRates["Low_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["High_Value", intRow].Value = txtHighValue.Text.Trim();
                                grdRates["High_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["Rate", intRow].Value = txtRate.Text.Trim();
                                grdRates["Rate", intRow].Style.BackColor = Color.Lavender;
                                //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                foreach (string s in lstTableColumns)
                                {
                                    if (dictGridValues[s] == grdRates[s, intRow].Value.ToString().Trim())
                                    {

                                    }
                                    else
                                    {
                                        //Write out to audit log
                                        writeAudit("RATES", "U - Update", s, dictGridValues[s], grdRates[s, intRow].Value.ToString().Trim());

                                    }

                                }

                            }
                            else
                            {
                                MessageBox.Show("Invalid Password.", "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion 

            }
        }

        private void writeAudit(string tablename, string function, string fieldname, string oldValue, string newValue)
        {
            string PK = string.Empty;
            foreach (string key in dictPrimaryKeyValues.Keys)
            {
                PK = PK + "<" + key.Trim() + "=" + dictPrimaryKeyValues[key] + ">";
            }

            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "AUDIT");
            audit.Clear();

            DataRow dr = audit.NewRow();
            dr["Type"] = function.Substring(0, 1);
            dr["TableName"] = tablename;
            dr["PK"] = PK;
            dr["FieldName"] = fieldname;
            dr["OldValue"] = oldValue;
            dr["NewValue"] = newValue;
            dr["UpdateDate"] = DateTime.Today.ToLongDateString();
            dr["UserName"] = BusinessLanguage.Userid;

            audit.Rows.Add(dr);
            audit.AcceptChanges();

            TB.saveCalculations2(audit, Base.DBConnectionString, " where type = 'x'", "AUDIT");
        }

        private void clearAllCalcValues(string _Tablename, string _Section)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Update " + _Tablename + " set ");
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, _Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                sb.Append(row["CALC_NAME"].ToString() + " = '0',");
            }

            if (sb.Length > 25)
            {
                sb.Append(strWhere);

                string strTemp = Convert.ToString(sb.Replace(",where", " Where"));
                TB.InsertData(Base.DBConnectionString, strTemp);
            }
        }

        private void btnProcessAll_Click(object sender, EventArgs e)
        {

            int intCheckLocks = checkLockInputProcesses();
            if (intCheckLocks == 0)
            {
                openTab(tabProcess);

                //checkProcess();
                calcCrewsandGangs();
            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }

        }

        private void deleteAllColumns(string Tablename)
        {
            //xxxxxxxxxxxxxxxxxxx
            //Create the earnings table schema
            createTheFile(Tablename);

            //Add the calculation columns.
            createEarningsColumns(Tablename);

            List<string> lstColumnNames = new List<string>();

            //extract the latest data from the base file e.g. Ganglink, Bonusshifts and replace data in the earningsfile.
            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                           "where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period.Trim() + "'");

            //Give the tempory file a name
            tb.TableName = Tablename + "EARN" + BusinessLanguage.Period.Trim();

            if (Tablename.ToUpper() == "BONUSSHIFTS")
            {
                #region Remove columns starting with DAY from BONUSSHIFTS
                //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                        lstColumnNames.Add(dc.ColumnName.Trim());
                    }
                    else
                    {

                    }
                }

                foreach (string s in lstColumnNames)
                {
                    tb.Columns.Remove(s);
                    tb.AcceptChanges();
                }

                lstColumnNames.Clear();
                #endregion
            }

            //Save the data to be processed to the earnings table.
            TB.saveCalculations2(tb, Base.DBConnectionString, " where section = '" + txtSelectedSection.Text.Trim() +
                                 "' or period != '" + BusinessLanguage.Period.Trim() + "'",
                                 tb.TableName.Trim());
            Application.DoEvents();
            //}
        }


        private void createTheFile(string Tablename)
        {
            //Check if earningstable exist - e.g. GangLinkEarn201108....if not...CREATE the table
            List<string> lstColumnNames = new List<string>();

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period.Trim());

            if (intCount > 0)
            {
            }
            else
            {
                //CREATE the earnings table:  GanglinkEarn201108
                //Extract the table into a temp file from the datafile e.g. GANGLINK, BONUSSHIFTS, DRILLERS etc.

                DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                               " where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

                //Give the tempory file a name
                tb.TableName = Tablename + "Earn" + BusinessLanguage.Period.Trim();

                #region remove day columns from bonusshifts
                if (Tablename.ToUpper() == "BONUSSHIFTS")
                {
                    #region Remove columns starting with DAY from BONUSSHIFTS
                    //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                    foreach (DataColumn dc in tb.Columns)
                    {
                        if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                        {
                            lstColumnNames.Add(dc.ColumnName.Trim());
                        }
                        else
                        {

                        }
                    }

                    foreach (string s in lstColumnNames)
                    {
                        tb.Columns.Remove(s);
                        tb.AcceptChanges();
                    }

                    lstColumnNames.Clear();
                    #endregion
                }

                strSqlAlter.Remove(0, strSqlAlter.Length);

                #endregion

                //First create the base table.  Why, because all these columns should be NOT NULL.  
                //The Formulas SHOULD be NULL when created

                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                    }
                    else
                    {
                        lstColumnNames.Add(dc.ColumnName);
                    }
                }

                //Create the earningstable e.g. BONUSSHIFTSEARN201108T

                TB.createEarningsTable(Base.DBConnectionString, tb.TableName, Tablename, lstColumnNames);

            }
        }

        private void createEarningsColumns(string Tablename)
        {
            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period);

            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName + BusinessLanguage.Period,
                                      Tablename + "EARN", Base.AnalysisConnectionString);

            foreach (DataRow row in tableformulas.Rows)
            {
                if (tb.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                }
                else
                {
                    strSqlAlter = strSqlAlter.Append(" ; Alter table " + Tablename + "EARN" + BusinessLanguage.Period + " add " +
                                                     row["CALC_NAME"].ToString().Trim() + " varchar(50) NULL");
                }
            }

            if (strSqlAlter.ToString().Trim().Length > 0)
            {
                StringBuilder bld = new StringBuilder();
                bld.Append("BEGIN transaction;" + strSqlAlter.ToString().Substring(1).Trim() + ";COMMIT transaction;");

                TB.InsertData(Base.DBConnectionString, bld.ToString().Trim());
                Application.DoEvents();
            }
            else
            {
            }
        }

        private void deleteAllCalcColumns(string Tablename)
        {
            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
            }
        }

        private void deleteAllCalcColumns(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
                }

            }
        }

        private void deleteAllCalcColumnsFromTempTable(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    Table.Columns.Remove(row["CALC_NAME"].ToString().Trim());                
                }
            }

            Table.AcceptChanges();
        }

        private void Calcs(string tablename, string phasename, string Delete)
        {
            if (Delete == "Y")
            {
                deleteAllColumns(tablename);
            }

            TB.insertProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period, tablename + "EARN", phasename, txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "N", "N", (string)DateTime.Now.ToLongTimeString(), Convert.ToString(++intProcessCounter));

        }

        private void openTab(TabPage tp)
        {
            this.tabInfo.SelectedTab = tp;

            Application.DoEvents();

        }

        private void calcCrewsandGangs()
        {
            string strTableName = "";

            for (int i = 1; i <= 4; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        //btnPhase1.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 1", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink10", "Y");
                        break;

                    case 2:
                        //btnPhase1.BackColor = Color.LightGreen;
                        //btnPhase2.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 2", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink20", "Y");
                        break;

                    case 3:
                        //btnPhase2.BackColor = Color.LightGreen;
                        //btnPhase3.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 3", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink30", "Y");
                        break;

                    case 4:
                        //btnPhase3.BackColor = Color.LightGreen;
                        //btnPhase4.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 4", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Header", "Base Calc Process", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());

                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink40", "Y");
                        break;
                }

                //executeFormulas(strTableName);
            }


            //btnPhase4.BackColor = Color.LightGreen;
            //Application.DoEvents();

        }

        private void calcCrewsandGangs(int counter)
        {
            string strTableName = "";

            for (int i = counter; i <= counter; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        Calcs("GangLink", "Ganglink10", "Y");
                        break;

                    case 2:
                        Calcs("GangLink", "Ganglink20", "N");
                        break;

                    case 3:
                        Calcs("GangLink", "Ganglink30", "N");
                        break;

                    case 4:
                        Calcs("GangLink", "Ganglink40", "N");
                        break;
                }


            }



            Application.DoEvents();

        }

        private void executeCostSheetFormulas(string TableName)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);
            string strprevPeriod = TableName;
            strSQL = "BEGIN transaction; insert into monitor values('" + Base.DBName + "','" + strprevPeriod + "','N','0','" + txtSelectedSection.Text.Trim() + "','0','0'); commit transaction; ";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        #region Open Tabs

        private void btnLockCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

         

        private void btnLockBonusShifts_Click(object sender, EventArgs e)
        {
            openTab(tabLabour);
        }


        #endregion

        private void btnCrewLevel_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                calcCrewsandGangs();

                evaluateStatus();
            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
        }

        private void btnEmplTeamCalcHeader_Click(object sender, EventArgs e)
        {
            evaluateStatus();
        }

        private void saveXXXTeamShifts(DataTable TeamShifts)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in TeamShifts.Rows)
            {

                strSQL.Append("insert into TeamShifts values('" + rr["SECTION"].ToString().Trim() +
                              "','" + rr["CONTRACT"].ToString().Trim() + "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["WAGECODE"].ToString().Trim() + "','" + rr["LINERESPCODE"].ToString().Trim() + "','" +
                              rr["EMPLOYEE_NO"].ToString().Trim() + "','" + rr["INITIALS"].ToString().Trim() + "','" +
                              rr["SURNAME"].ToString().Trim() + "','" + rr["REGISTER"].ToString().Trim() + "','" +
                              rr["DATEFROM"].ToString().Trim() + "','" + rr["EMPLOYEEPRODUCTIONBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEDRESSINGBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEAWOPPENALTYBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEAWOPDRESSNGPENALTYBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEHYDROBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEESTOPEPROCESSBONUS"].ToString().Trim() + "')");



            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void saveXXXTeamProd(DataTable Teamprod)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in Teamprod.Rows)
            {
                //"CREATE TABLE TEAMPROD (SECTION char(50), CONTRACT Char(50), WORKPLACE Char(50), " +
                //    "GANG Char(50),WPNAME Char(50),WPSHIFTS Char(50),WPSHIFTSTOTAL Char(50), WPSQM Char(50), " +
                //    "WPFOOTWALL Char(50),WPSTOPEWIDTH Char(50),WPSTOPEWIDTHRATE Char(50), WPSTOPEWIDTHBONUS Char(50), " +
                //    "WPCONTRACTORBONUS Char(50),WPTOTALBONUS Char(50))";

                strSQL.Append("insert into TeamProd values('" + rr["SECTION"].ToString().Trim() + "','" + rr["CONTRACT"].ToString().Trim() +
                              "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["CREWNO"].ToString().Trim() + "','" + rr["WPNAME"].ToString().Trim() + "','" + rr["WPSHIFTS"].ToString().Trim() +
                              "','" + rr["WPSHIFTSTOTAL"].ToString().Trim() + "','" + rr["WPSQM"].ToString().Trim() + "','" +
                              rr["WPFOOTWALL"].ToString().Trim() + "','" + rr["WPSTOPEWIDTH"].ToString().Trim() +
                              "','" + rr["WPSTOPEWIDTHRATE"].ToString().Trim() + "','" + rr["WPSTOPEWIDTHBONUS"].ToString().Trim() +
                              "','" + rr["WPCONTRACTORBONUS"].ToString().Trim() + "','" + rr["WPTOTALBONUS"].ToString().Trim() + "');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void WriteCSV(DataTable dt)
        {
            StreamWriter sw;
            string filePath = strServerPath + ":\\Crew.csv";


            sw = File.CreateText(filePath);

            try
            {
                // write the data in each row & column
                int intcounter = 0;
                foreach (DataRow row in dt.Rows)
                {
                    // recreate an empty Stringbuilder through each row iteration.
                    StringBuilder rowToWrite = new StringBuilder();

                    for (int counter = 0; counter <= dt.Columns.Count - 1; counter++)
                    {
                        if (intcounter == 0)
                        {
                            foreach (DataColumn column in dt.Columns)
                            {
                                //rowToWrite.Append("'" + column.ColumnName + "'");
                                rowToWrite.Append("'" + column.ColumnName + "'");
                            }
                            rowToWrite.Replace("''", "','");
                            rowToWrite.Replace("'", "");

                            rowToWrite.Append("\r\n");
                            sw.Write(rowToWrite);
                            rowToWrite.Remove(0, rowToWrite.Length);
                        }
                        intcounter = intcounter + 1;
                        rowToWrite.Append("'" + row[counter] + "'");
                    }

                    rowToWrite.Replace("''", "','");
                    rowToWrite.Replace("'", "");

                    rowToWrite.Append("\r\n");
                    sw.Write(rowToWrite);
                }
            }
            catch
            {
                //("An error occurred while attempting to build the CSV file. " + e.Message);
            }
            finally
            {
                sw.Close();
            }
        }


        private void btnDeleteRow_Click_1(object sender, EventArgs e)
        {

            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {
               

                case "tabEmplPen":
                    #region tabEmployeePenalty

                    intRow = grdEmplPen.CurrentCell.RowIndex;
                    intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                    if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from EmployeePenalties " +
                                 " Where Section = '" + grdEmplPen["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdEmplPen["Period", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdEmplPen["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and PenaltyInd = '" + grdEmplPen["PenaltyInd", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateEmployeePenalties();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabOfficials":
                    #region tabOfficials

                    intRow = grdOfficials.CurrentCell.RowIndex;
                    intColumn = grdOfficials.CurrentCell.ColumnIndex;

                    if (grdOfficials["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from Officials " +
                                 " Where Section = '" + grdOfficials["Section", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdOfficials["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and Gang = '" + grdOfficials["Gang", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdOfficials["Period", intRow].Value.ToString().Trim() +
                                 "' and Wagecode = '" + grdOfficials["Wagecode", intRow].Value.ToString().Trim() +
                                 "' and ProcessCode = '" + grdOfficials["ProcessCode", intRow].Value.ToString().Trim() +
                                 "' and Linerespcode = '" + grdOfficials["Linerespcode", intRow].Value.ToString().Trim() +
                                 "' and Designation = '" + grdOfficials["Designation", intRow].Value.ToString().Trim() +
                                 "' and Designation_Desc = '" + grdOfficials["Designation_Desc", intRow].Value.ToString().Trim() + 
                                 "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateOfficials();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabParticipants":
                    #region tabParticipants

                    intRow = grdParticipants.CurrentCell.RowIndex;
                    intColumn = grdParticipants.CurrentCell.ColumnIndex;

                    if (grdParticipants["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from Participants " +
                                 " Where Section = '" + grdParticipants["Section", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdParticipants["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and Employee_Name = '" + grdParticipants["Employee_name", intRow].Value.ToString().Trim() +
                                 "' and hod = '" + grdParticipants["HOD", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdParticipants["PERIOD", intRow].Value.ToString().Trim() +
                                 "' and Designation = '" + grdParticipants["Designation", intRow].Value.ToString().Trim() +
                                 "' and Designation_Desc = '" + grdParticipants["Designation_Desc", intRow].Value.ToString().Trim() + 
                                 "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateParticipants();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion
            }
        }

        protected virtual void FrontDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteFullBeginTag("HTML");
            writer.WriteFullBeginTag("Head");
            writer.RenderBeginTag(System.Web.UI.HtmlTextWriterTag.Style);
            writer.Write("<!--");

            StreamReader sr = File.OpenText(strServerPath + ":\\koos.html");
            String input;
            while ((input = sr.ReadLine()) != null)
            {
                writer.WriteLine(input);
            }
            sr.Close();
            writer.Write("-->");
            writer.RenderEndTag();
            writer.WriteEndTag("Head");
            writer.WriteFullBeginTag("Body");
        }

        protected virtual void RearDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteEndTag("Body");
            writer.WriteEndTag("HTML");
        }

        private void printHTML(DataTable dt, string TabName)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = "c:\\icalc\\koos.html";
                try
                {

                    StreamWriter SW = new StreamWriter(OPath);
                    //StringWriter SW = new StringWriter();
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {
                             
                            HTMLWriter.WriteLine("HARMONY - Phakisa Mine - " + TabName + " - " + BusinessLanguage.Period);
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("==============================");
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();

                            grid.RenderControl(HTMLWriter);
                            //RearDecorator(HTMLWriter);

                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();


                    System.Diagnostics.Process P = new System.Diagnostics.Process();
                    P.StartInfo.WorkingDirectory = strServerPath + ":\\Program Files\\Internet Explorer";
                    P.StartInfo.FileName = "IExplore.exe";
                    P.StartInfo.Arguments = "C:\\icalc\\koos.html";
                    P.Start();
                    P.WaitForExit();


                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void btnLoad_Click_1(object sender, EventArgs e)
        {
            if (listBox3.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select the number of measuring shifts", "Information", MessageBoxButtons.OK);
            }
            else
            {
                if (txtSelectedSection.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Please select a section and the correct month measuring shifts for the section.", "Information", MessageBoxButtons.OK);
                }
                else
                {
                    int intCalendarProcesses = checkLockCalendarProcesses();

                    if (intCalendarProcesses == 0)
                    {
                        //The calendar is currently locked and has to be unlocked before any updates can be doen
                        MessageBox.Show("Please UNLOCK Calendar before changing any section's' shifts.");
                    }
                    else
                    {
                       //string selectedSection = txtSelectedSection.Text.Trim();
                        string grdSection = grdCalendar["SECTION", intFiller].Value.ToString().Trim();
                        if (label36.Text.Trim() == grdSection)
                        {
                            Base.updateCalendarRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                             txtBonusType.Text.Trim(), label36.Text.Trim(),
                                                             txtPeriod.Text.ToString().Trim(),
                                                             (Convert.ToDateTime(dateTimePicker1.Text)).ToString("yyyy-MM-dd"),
                                                             (Convert.ToDateTime(dateTimePicker2.Text)).ToString("yyyy-MM-dd"),
                                                             listBox3.SelectedItem.ToString().Trim());
                            Application.DoEvents();

                            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", 
                                                                             " Where period = '" + BusinessLanguage.Period + "'");
                            if (Calendar.Rows.Count > 0)
                            {
                                DateTime EarliestStart = Convert.ToDateTime(Calendar.Rows[0]["FSH"].ToString().Trim());
                                DateTime LatestEnd = Convert.ToDateTime(Calendar.Rows[0]["LSH"].ToString().Trim());
                                int j = 0;
                                for (int i = 0; i <= Calendar.Rows.Count - 1; i++)
                                {

                                    if (Convert.ToDateTime(Calendar.Rows[i]["FSH"].ToString().Trim()) < EarliestStart &&
                                        Calendar.Rows[i]["SECTION"].ToString().Trim() != "OFF")
                                    {
                                        EarliestStart = Convert.ToDateTime(Calendar.Rows[i]["FSH"].ToString().Trim());
                                    }
                                    if (Convert.ToDateTime(Calendar.Rows[i]["LSH"].ToString().Trim()) > LatestEnd &&
                                        Calendar.Rows[i]["SECTION"].ToString().Trim() != "OFF")
                                    {
                                        LatestEnd = Convert.ToDateTime(Calendar.Rows[i]["LSH"].ToString().Trim());
                                    }

                                    if (Calendar.Rows[i]["SECTION"].ToString().Trim() == "OFF")
                                    {
                                        j = i;
                                    }

                                }

                                Calendar.Rows[j]["FSH"] = Convert.ToDateTime(EarliestStart).ToString("yyyy-MM-dd");
                                Calendar.Rows[j]["LSH"] = Convert.ToDateTime(LatestEnd).ToString("yyyy-MM-dd");

                                TB.saveCalculations2(Calendar, Base.DBConnectionString, " where period = '" + BusinessLanguage.Period + "'", "Calendar");
                                evaluateCalendar();
                            }
                        }

                        else
                        {
                            MessageBox.Show("Selected section not the same as grid section.", "Informations", MessageBoxButtons.OK);
                        }

                        //Extract Calendar again and insert into 

                        grdCalendar.DataSource = Calendar;
                    }
                }
            }
        }


        private void label83_Click(object sender, EventArgs e)
        {
            extractDBTableNames(listBox1);
        }

        private void btnLockPaysend_Click(object sender, EventArgs e)
        {
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();
            
         
        }

        private void btnEmployeeCostsheet_Click(object sender, EventArgs e)
        {
            Calcs("Miners", "Miners", "N");
        }

        private void btnPrint_Click_1(object sender, EventArgs e)
        {
            switch (tabInfo.SelectedTab.Name)
            {
                case "tabMiners":
                    #region tabMiners

                    DataTable dt = Base.extractPrintData(Base.DBConnectionString, "Miners", strWhere);
                    deleteAllCalcColumns("Miners", dt);
                    if (dt.Rows.Count > 0)
                    {
                        dt.Columns.Remove("BUSSUNIT");
                        dt.Columns.Remove("MININGTYPE");
                        dt.Columns.Remove("BONUSTYPE");
                        dt.Columns.Remove("SAFETYIND");
                        dt.Columns.Remove("SHIFTS_WORKED");
                        dt.AcceptChanges();

                        printHTML(dt, "Miners");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabParticipants":
                    #region tabOfficials

                    dt = Base.extractPrintData(Base.DBConnectionString, "Participants", "");
                    deleteAllCalcColumns("Participants", dt);
                    if (dt.Rows.Count > 0)
                    {

                        dt.AcceptChanges();

                        printHTML(dt, "Participants");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabGangLinking":
                    #region tabGangLinking

                    dt = Base.extractPrintData(Base.DBConnectionString, "GangLink", strWhere);
                    deleteAllCalcColumns("GangLink", dt);
                    dt.AcceptChanges();
                    if (dt.Rows.Count > 0)
                    {

                        printHTML(dt, "GangLinking");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabAbnormal":
                    #region tabAbnormal

                    dt = Base.extractPrintData(Base.DBConnectionString, "Abnormal", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Abnormal");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabLabour":
                    #region tabLabour

                    dt = Base.extractPrintData(Base.DBConnectionString, "BonusShifts", strWhere);
                    deleteAllCalcColumns("BonusShifts", dt);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "BonusShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabSurvey":
                    #region tabSurvey

                    dt = Base.extractPrintData(Base.DBConnectionString, "Survey", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Survey");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    dt = Base.extractPrintData(Base.DBConnectionString, "EmployeePenalties", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "EmployeePenalties");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }
                    break;
                    #endregion

                case "tabOffday":
                    #region tabOffdays

                    dt = Base.extractPrintData(Base.DBConnectionString, "Offdays", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Offdays");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabCalendar":
                    #region tabCalendar

                    dt = Base.extractPrintData(Base.DBConnectionString, "Calendar", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Calendar");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabClockShifts":
                    #region tabClockShifts

                    dt = Base.extractPrintData(Base.DBConnectionString, "ClockedShifts", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "ClockedShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Rates", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Rates");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabMonitor":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Monitor", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Monitor");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

            }
        }

        private void calcStopeData()
        {
            Base.Period = txtPeriod.Text.Trim();
            //Base.Period = "200909";

            SqlConnection stopeConn = Base.StopeConnection;
            stopeConn.Open();

            try
            {
                DataTable ContractTotals = TB.getContractCrewOfficialBonus(Base.StopeConnectionString, "STOPING", txtSelectedSection.Text.Trim());

                stopeConn.Close();

                TB.updateDSShiftbossCrewBonus(Base.DBConnectionString, ContractTotals);
            }
            catch { }


        }

        private void btnBaseCalcsHeader_Click(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text.Trim() != "***")
            {
                int intCheckLocks = checkLockInputProcesses();

                if (intCheckLocks == 0)
                {
                    this.Cursor = Cursors.WaitCursor;
                    DBDefault_Click_1("Method", null);
                    btnx.Visible = true;
                    btnx.Enabled = true;
                    btnx.Text = "Run";
                    TB.deleteProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period);
                    TB.deleteAllExcept(Base.DBConnectionString, "Monitor");
                    TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.MEASSECTION = PARTICIPANTS.section " +
                                     "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                     " AND OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period.Trim() + "'");
                    TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.DESIGNATION = PARTICIPANTS.DESIGNATION " +
                                     "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                     " AND OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period.Trim() + "'");
                    TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.DESIGNATION_DESC = PARTICIPANTS.DESIGNATION_DESC " +
                                     "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                     " AND OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period.Trim() + "'");
                    Calcs("Officials", "Officialsearn5", "Y");
                    Calcs("Officials", "Officialsearn10", "N");
                    Calcs("Officials", "Officialsearn15", "N");
                    Calcs("Officials", "Officialsearn18", "N");
                    Calcs("Officials", "Officialsearn20", "N");
                    Calcs("Officials", "Officialsearn25", "N");
                    Calcs("Officials", "Officialsearn30", "N");
                    Calcs("Officials", "Officialsearn40", "N");
                    Calcs("Exit", "Exit", "N");

                    btnBaseCalcs.BackColor = Color.Orange;
                    btnOfficialsCalcs.BackColor = Color.Orange;

                    TB.updateStatusFromArchive(Base.DBConnectionString, "N", "OFFICIALSEARN10", "OFF");
                    TB.updateStatusFromArchive(Base.DBConnectionString, "N", "OFFICIALSEARN40", "OFF");
                    TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Exit", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
                    //Base.backupDatabase3(Base.DBConnectionString, Base.DBName, Base.BackupPath);
                    this.Cursor = Cursors.Arrow;
                }

                else
                {
                    MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Please select a section", "Informations", MessageBoxButtons.OK);
            }
           

        }

        private void btnBaseCalcs_Click(object sender, EventArgs e)
        {
           
        }


        private void grdActiveSheet_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int columnnr = e.ColumnIndex;
                DialogResult result = MessageBox.Show("Do you want to delete the column:  " + grdActiveSheet.Columns[columnnr].HeaderText + "?", "INFORMATION", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    //int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
                    TB.removeColumn(Base.DBConnectionString, TB.TBName, grdActiveSheet.Columns[columnnr].HeaderText);
                    DoDataExtract("");
                    grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);
                }
                else
                {
                    if (listBox1.SelectedItem.ToString().Trim() == "MONITOR")
                    {

                        string strSQL = "Begin transaction; Delete from monitor; commit transaction";
                        TB.InsertData(Base.DBConnectionString, strSQL);
                        Application.DoEvents();

                    }
                }
            }

            else
            {
                AConn = Analysis.AnalysisConnection;
                AConn.Open();
                string FormulaTableName = string.Empty;
                if (TB.TBName.Trim().ToUpper().Contains("EARN") && TB.TBName.Trim().ToUpper().Contains("20"))
                {
                    FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));   //xxxxxxxxxxxxxxxxxx
                }
                else
                {
                    FormulaTableName = TB.TBName;
                }

                DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName, Base.AnalysisConnectionString);

                foreach (DataRow dt in tempDataTable.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void grdCalendar_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["LSH"].ToString().Trim());
                label36.Text = Calendar.Rows[e.RowIndex]["SECTION"].ToString().Trim();
                label36.Visible = true;
                intFiller = e.RowIndex;
                }
        }
 

        private void grdRates_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if (grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;

                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                    btnInsertRow.Enabled = true;
                }

                txtRateType.Text = grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim();
                txtLowValue.Text = grdRates["LOW_VALUE", e.RowIndex].Value.ToString().Trim();
                txtHighValue.Text = grdRates["HIGH_VALUE", e.RowIndex].Value.ToString().Trim();
                txtRate.Text = grdRates["RATE", e.RowIndex].Value.ToString().Trim();
            }
        }

    

        private void payrollSend_Click(object sender, EventArgs e)
        {
             
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();
            

        }

        

        private void dataFilter_Click(object sender, EventArgs e)
        {
            if (General.textTestSQL.ToString().Trim().Length > 0)
            {
                scrQuerySQL testsql = new scrQuerySQL();
                testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
                testsql.Show();
            }
            else
            {
                MessageBox.Show("No SQL to pass", "Information", MessageBoxButtons.OK);
            }
        }

        private void dataPrintTables_Click(object sender, EventArgs e)
        {

        }

        

        private void TBCreateSpreadsheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (openDialog.ShowDialog() != DialogResult.OK) return;
                //grpData.Enabled = false;
                string filename = openDialog.FileName;
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                fs.Close();

                if (spreadsheet.WorkbookData.Tables.Count > 0)
                {
                    switch (string.IsNullOrEmpty(Base.DBName))
                    {
                        case true:
                            MessageBox.Show("Create or select a database.", "DATABASE NEEDED!", MessageBoxButtons.OK);
                            break;

                        case false:
                            saveTheSpreadSheetToTheDatabase();
                            MessageBox.Show("Successfully Uploaded.", "Information", MessageBoxButtons.OK);
                            break;
                        default:

                            break;
                    }
                }

                 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read file: \n" + ex.Message);
            }
        }

        private void TBDeleteTable_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Delete table: " + TB.TBName + " ? ", "Confirm", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
                    extractDBTableNames(listBox1);
                    TB.deleteDataTableFromCollection(TB.DBName);
                    TB.TBName = "";
                    TBFormulas.Tablename = "";
                    loadInfo();
                    break;


                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteCalcColumns_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Confirm DELETE of calculated columns from table: " + TBFormulas.Tablename + "?", "", MessageBoxButtons.YesNo);

            switch (result1)
            {
                case DialogResult.Yes:

                    DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, TB.TBName, Base.AnalysisConnectionString);
                    foreach (DataRow row in tableformulas.Rows)
                    {
                        TB.removeColumn(Base.DBConnectionString, TB.TBName, row["CALC_NAME"].ToString());

                    }
                    loadInfo();
                    break;

                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteAllTables_Click(object sender, EventArgs e)
        {
            foreach (string s in listBox1.Items)
            {
                TB.TBName = s.Trim();
                bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
            }
            extractDBTableNames(listBox1);
            loadInfo();
        }

        private void DBCreate_Click(object sender, EventArgs e)
        {

        }

        private void createNewDatabase(string Databasename)
        {

        }

 
        private void evaluateStatusButtons()
        {
            btnInsertRow.Enabled = false;
            btnUpdate.Enabled = false;
            btnDeleteRow.Enabled = false;
            btnLoad.Enabled = false;
            btnPrint.Enabled = false;
            btnLock.Enabled = false;

            panelInsert.BackColor = Color.Cornsilk;
            panelUpdate.BackColor = Color.Cornsilk;
            panelDelete.BackColor = Color.Cornsilk;
            panelPreCalcReport.BackColor = Color.Cornsilk;
        }

        private void btnx_Click_1(object sender, EventArgs e)
        {

            btnx.Text = "Running";
            btnx.Enabled = false;
            btnRefresh.Visible = true;
            execute();
            //refreshExecution();

        }

           private void execute()
        {
           
            System.Diagnostics.Process P = new System.Diagnostics.Process();

            switch(BusinessLanguage.Env)
            {
                case "Production":
                    strName = "PhakisaOffSerP";
                    P.StartInfo.WorkingDirectory = @"z:\Harmony\Phakisa\Production\Core";
                    P.StartInfo.FileName = strName + ".exe";

                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;

                case "Test":
                    strName = "Archive2";
                    P.StartInfo.WorkingDirectory = "C:\\windows\\system32\\GM\\BACKUPS";
                    P.StartInfo.FileName = strName + ".exe" ;

                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;

                case "Development":

                    strName = "PhakisaOffSerD";
                    P.StartInfo.WorkingDirectory = @"C:\iCalc\Harmony\ServerProjects\Phakisa\Core\";
                    P.StartInfo.FileName = strName + ".exe";

                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;
            }           
        }

        private void btnRefresh_Click_1(object sender, EventArgs e)
        {
           
            evaluateStatus();
            evaluateStatusButtons();
            
        }

        private void btnTeamPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("PhakisaGN");
            this.Cursor = Cursors.Arrow;

        }

        private void printCostsheet_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("PhakisaMS");
            this.Cursor = Cursors.Arrow;
        }

        private void printAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.StartReport("STP_AUTHGSUM");
            MessageBox.Show("To be implemented", "Information", MessageBoxButtons.OK);
            this.Cursor = Cursors.Arrow;
        }

        private void btnCostsheetAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.StartReport("STPTM_CAS");
            MessageBox.Show("To be implemented", "Information", MessageBoxButtons.OK);
            this.Cursor = Cursors.Arrow;
        }

        private void TBExport_Click_1(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void btnChangePeriod_Click(object sender, EventArgs e)
        {
            //Gets the name of all open forms in application
            foreach (Form form in Application.OpenForms)
            {
                if (form is scrLogon)
                {
                    form.Show(); //Show the form
                    break;
                }
            }
            exitValue = 2;//Change exit value

            this.Close(); //Close the current window

        }

        private void scrOfficials_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (exitValue == 0)
            {
                DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        //this.Close();
                        //scrMain main = new scrMain();
                        //main.MainLoad(BusinessLanguage, DB, Survey, Labour, Miners, Designations, Occupations, Clocked, EmplList, EmplPen, Configs);
                        //main.ShowDialog();
                        myConn.Close();
                        AAConn.Close();
                        AConn.Close();
                        //this.Close();
                        exitValue = 1;
                        Application.Exit();
                        break;

                    case DialogResult.No:
                        e.Cancel = true;
                        break;
                }
                if (exitValue == 2)
                {
                    exitValue = 1;
                    this.Close();
                }
            }
        }


        private void btnAttendance_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            evaluateLabour();
            if (Labour.Rows.Count == 0)
            {
                MessageBox.Show("No Labour records to print for the section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
            else
            {
                DataTable temp = Labour.Copy();
                deleteAllCalcColumnsFromTempTable("BonusShifts", temp);
                temp.Columns.Remove("TEAMLEADERIND");
                TB.createAttendanceTable(Base.DBConnectionString, temp);

                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
                mm.Init(strMetaReportCode);
                mm.StartReport("STPTMTA");
            }
            this.Cursor = Cursors.Arrow;
        }

        private void btnSearchEmployNr_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = true;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            txtSearchEmplyNr.Focus();
        }

        private void btnEmployName_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = true;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            txtSearchEmplName.Focus();
        }

        private void btnSearchGang_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = true;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            txtSearchGang.Focus();
        }

        private void txtSearchEmplyNr_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            searchEmplNr = txtSearchEmplyNr.Text.ToString();
            searchEmplName = "";
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang,grdLabour); //Calls the metod

        }

        private void txtSearchEmplName_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = txtSearchEmplName.Text.ToString();
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang,grdLabour); //Calls the metod

        }

        private void txtSearchGang_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = "";
            searchEmplGang = txtSearchGang.Text.ToString();
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang,grdLabour); //Calls the metod
        }

        public void searchBonus(string nr, string name, string gang,DataGridView Grid)
        {
            //Sets the details passed to lower case
            nr = nr.ToLower();
            name = name.ToLower();
            gang = gang.ToLower();

            //Gets the length
            int nrLenght = nr.Length;
            int nameLenght = name.Length;
            int gangLenght = gang.Length;

            // Ensuring length are always 1 and not 0 as
            // "" can not be tested.
            if (nrLenght == 0)
            {
                nrLenght = 1;
            }
            if (nameLenght == 0)
            {
                nameLenght = 1;
            }
            if (gangLenght == 0)
            {
                gangLenght = 1;
            }

            //Iterate through all the rows in the grid
            for (int i = 0; i < Grid.Rows.Count - 1; i++)
            {
                //Gets the values of the grid in the different columns
                string nrColumn = Grid.Rows[i].Cells["Employee_No"].Value.ToString();  //Cells from grid count from left starting at 0
                string nameColumn = Grid.Rows[i].Cells[1].Value.ToString();
                string gangColumn = Grid.Rows[i].Cells["Gang"].Value.ToString();

                //Sets the values from grid to lowercase for testing
                nrColumn = nrColumn.ToLower();
                nameColumn = nameColumn.ToLower();
                gangColumn = gangColumn.ToLower();

                //Gets the same amount from the grid string as was entertered bty the user to 
                //ensure the string can be tested
                nrColumn = nrColumn.Substring(1, nrLenght);//Start at 1 to throw away the aphabetic nr
                nameColumn = nameColumn.Substring(0, nameLenght);
                gangColumn = gangColumn.Substring(0, gangLenght);

                //Compares the different strings
                if (nr == nrColumn) //Employee nr
                {
                    //Empty the string not used
                    nameColumn = "";
                    gangColumn = "";
                    Grid.ClearSelection(); // Clears all past selection
                    Grid.Rows[i].Selected = true; //Selects the current row
                    Grid.FirstDisplayedScrollingRowIndex = i; //Jumps automatically to the row
                    break; //breaks the loop
                }
               
                if (gang == gangColumn) //Gang
                {
                    nrColumn = "";
                    nameColumn = "";
                    Grid.ClearSelection();
                    Grid.Rows[i].Selected = true;
                    Grid.FirstDisplayedScrollingRowIndex = i;
                    break;
                }
            }
        }

        

        private void dataPrintFormulas_Click(object sender, EventArgs e)
        {
            DataTable dt = Base.dataPrintFormulasBonusShifts(Base.AnalysisConnectionString, Base.DBName, "OFFICIALS");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row["FORMULA_NAME"].ToString().Trim().Length > 3 && row["FORMULA_NAME"].ToString().Trim().Substring(0, 3) == "SQL")
                    {
                    }
                    else
                    {
                        switch (row["INPUTORDER"].ToString().Trim())
                        {
                            case "0":
                                row["INPUTORDER"] = "A = ";
                                break;
                            case "1":
                                row["INPUTORDER"] = "B = ";
                                break;
                            case "2":
                                row["INPUTORDER"] = "C = ";
                                break;
                            case "3":
                                row["INPUTORDER"] = "D = ";
                                break;
                            case "4":
                                row["INPUTORDER"] = "E = ";
                                break;
                            case "5":
                                row["INPUTORDER"] = "F = ";
                                break;
                            case "6":
                                row["INPUTORDER"] = "G = ";
                                break;
                            case "7":
                                row["INPUTORDER"] = "H = ";
                                break;
                            case "8":
                                row["INPUTORDER"] = "I = ";
                                break;
                            case "9":
                                row["INPUTORDER"] = "J = ";
                                break;
                        }
                    }
                }
                printHTML(dt, "OFFICIALS");
            }
            else
            {
                MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
            }
        }

        private void auditByTable_Click(object sender, EventArgs e)
        {
            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Audit", " where tablename = 'Ganglink'");
            string[] auditcolumns = new string[10];

            string test = audit.Rows[0]["PK"].ToString().Trim();
            int testlength = test.Length;

            for (int i = 0; i <= 9; i++)
            {
                int tstLength = test.IndexOf(">");
                if (tstLength != -1)
                {
                auditcolumns[i] = test.Substring(0, tstLength).Replace("<", "").Trim();
                test = test.Substring(test.IndexOf(">") + 1);
                }

            }





        }

        private void btnEmplyeAudit_Click(object sender, EventArgs e)
        {


            #region extract the sheet name and FSH and LSH of the extract
            string FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\ADTeam_201004.xls";
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);

            #region remove invalid records
            // Delete records that does not conform to configurations
            //foreach (DataRow row in dt.Rows)
            //{
            //    if ((row["GANG NAME"].ToString().Substring(5, 1) == "A" || row["GANG NAME"].ToString().Substring(5, 1) == "B" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "C" || row["GANG NAME"].ToString().Substring(5, 1) == "D" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "E" || row["WAGE CODE"].ToString() == "245M003" ||
            //        row["WAGE CODE"].ToString() == "400M009" || row["WAGE CODE"].ToString() == "245M001" ||
            //        row["WAGE CODE"].ToString() == "246M004" || row["WAGE CODE"].ToString() == "400M009")
            //        && (row["GANG NAME"].ToString().Substring(0, 5) == txtSelectedSection.Text.Trim()))
            //    {
            //    }
            //    else
            //    {
            //        //row.Delete();
            //    }

            //}

            //dt.AcceptChanges();

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }

            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);
            noOFDay = intNoOfDays;

            if (intNoOfDays <= 44)
            {
                for (int j = intNoOfDays + 1; j <= 44; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;


            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "OFFICIALS";
                if (row["GANG"].ToString().Length > 0)
                {
                    row["SECTION"] = row["GANG"].ToString().Substring(0, 5);
                }
                else
                {
                    row["SECTION"] = "XXX";
                }

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion

            //Write to the database
            // TB.saveCalculations2(dt, Base.DBConnectionString, strWhere, "CLOCKEDSHIFTS");

            // Application.DoEvents();

            // grdClocked.DataSource = dt;
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            //string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
            //                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where (section = '"
            //                + txtSelectedSection.Text.Trim() + "' or WAGE_DESCRIPTION = 'STOPER')";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

            // BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);  

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;
            sheetlhs = SheetLSH;
            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                // DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

                #region Extract the ganglinking of the current section
                ////Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                ////before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                ////This is done in the methord extractGangLink

                //DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GANGLINK", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current ganglinking for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractGangLink();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractGangLink();
                //}

                //cboMinersGangNo.Items.Clear();
                //lstNames = TB.loadDistinctValuesFromColumn(Labour, "Gang");
                //if (lstNames.Count > 1)
                //{

                //    foreach (string s in lstNames)
                //    {
                //        if (cboMinersGangNo.Items.Contains(s))
                //        { }
                //        else
                //        {
                //            cboMinersGangNo.Items.Add(s.Trim());
                //        }
                //    }
                //}

                #endregion

                #region Extract the miners of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the method extractMiners

                //temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MINERS", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current MINERS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractMiners();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractMiners();
                //}
                #endregion

                fillFixTable(fixShifts, sheetfhs, sheetlhs, intNoOfDays, intStartDay, intStopDay);
                this.Cursor = Cursors.Arrow;
                //}
            }

        }

        public void fillFixTable(DataTable clockedTable, DateTime SheetFSH, DateTime SheetLSH, int intNoOfDays, int DayStart, int DayEnd)
        {
            //Calculate the shifts in the clockedshifts table and insert all in a fixed
            //table that cannot be changed by the user!

            string SQLTable = "IF OBJECT_ID(N'emplshiftfix', N'U')IS NOT NULL DROP TABLE EMPLSHIFTFIX create table EMPLSHIFTFIX (employeeno char(20),shiftsfix char(20)) truncate table EMPLSHIFTFIX";
            Base.VoidQuery(Base.DBConnectionString, SQLTable);

            #region Calculate the shifts per employee en output to bonusshifts

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                int intSubstringLength = 0;
                int intShiftsWorked = 0;
                int intAwopShifts = 0;
                int shiftsCheck = 0;
                StringBuilder sqlInsertFixShifts = new StringBuilder("BEGIN TRANSACTION; ");

                foreach (DataRow row in clockedTable.Rows)
                {
                    foreach (DataColumn column in clockedTable.Columns)
                    {
                        if ((column.ColumnName.Substring(0, 3) == "DAY"))
                        {

                            if (column.ColumnName.ToString().Length == 4)
                            {
                                intSubstringLength = 1;
                            }
                            else
                            {
                                intSubstringLength = 2;
                            }

                            if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                               Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                            {

                                if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u")
                                {
                                    intShiftsWorked = intShiftsWorked + 1;
                                    shiftsCheck = 1;
                                }
                                else
                                {
                                    if (row[column].ToString().Trim() == "A")
                                    {
                                        intAwopShifts = intAwopShifts + 1;
                                    }
                                    else { }

                                }
                            }
                            else
                            {
                                row[column] = "*";
                            }
                        }
                        else
                        {
                            if (column.ColumnName == "BONUSTYPE")
                            {
                                row["BONUSTYPE"] = "TEAM";
                            }
                        }
                    }//foreach datacolumn

                    row["SHIFTS_WORKED"] = intShiftsWorked;

                    string emplNr = row["employee_no"].ToString();
                    workedShiftsFixedClockedShift = intShiftsWorked;
                    intShiftsWorked = 0;
                    intAwopShifts = 0;
                    if (shiftsCheck == 1)
                    {
                        sqlInsertFixShifts.Append("INSERT INTO EMPLSHIFTFIX VALUES ('" + emplNr.Trim() + "','" + workedShiftsFixedClockedShift.ToString().Trim() + "');");
                    }
                }

                sqlInsertFixShifts.Append(" COMMIT TRANSACTION");


                Base.VoidQuery(Base.DBConnectionString, sqlInsertFixShifts.ToString());

                //DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

            }
        }

        private void lstBErrorLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedNr = lstBErrorLog.SelectedItem.ToString();
            if (selectedNr != "Employee Nr")
            {

                selectedNr = selectedNr.Remove(0, 1);
                int last = selectedNr.LastIndexOf("-");
                selectedNr = selectedNr.Remove(last - 1).Trim();
                txtSearchEmplyNr.Visible = true;
                txtSearchEmplyNr.Text = selectedNr;
            }
        }

        private void hideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grdActiveSheet.Columns[columnnr].Visible = false;
            

        }

       
        public void changeRights()
        {
            InputBoxResult pass = InputBox.Show("Password: ");

            string paas = pass.Text.ToString();
            if (pass.ReturnCode == DialogResult.OK)
            {
               
                if (paas == "admin")
                {
                     txtShifts.Enabled = true; 
                    
                }
                else
                {
                    MessageBox.Show("You do not have the right to change this box");
                    txtShifts.Enabled = false; 
                    txtShifts.Text = ""; 
                    
                }
            }

        }

       

        private void txtPayShifts_Click(object sender, EventArgs e)
        {
            changeRights();

        }

        private void txtShifts_Click(object sender, EventArgs e)
        {
            changeRights();
        }

        private void btnSurveySummary_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("SurveySumSTM");
            this.Cursor = Cursors.Arrow;
        }

      
 
        private void calcTime_Tick(object sender, EventArgs e)
        {
            btnRefresh_Click_1("Method", null);
        }

        private bool createZipFolder(string path, string databasename)
        {
            path = Base.BackupPath.Replace(Base.BackupPath.Substring(0, 2), "C:") + "\\" + databasename + DateTime.Today.ToString("yyyyMMdd");
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void FastZipCompress(string pathDBBackup)
        {
            //FastZip fZip = new FastZip();

            //fZip.CreateZip("C:\\ZipTest\\test.zip", pathDBBackup, false, ".bak$");

        }

        private static void FastZipCompress(string pathDBBackup, string zipname)
        {
            FastZip fZip = new FastZip();

            fZip.CreateZip("C:\\icalc\\" + zipname + ".zip", pathDBBackup.Replace("xxx.bak", ""), false, ".bak$");

        }

        private void BackupDB(string connectionstring, string dbname, string dbPath)
        {
            Cursor.Current = Cursors.Arrow;
            bool check = false;
            check = Base.backupDatabase3(connectionstring, dbname, dbPath);

            //Copy the file to the C:\drive
            if (check == true)
            {
                //MessageBox.Show("Source = " + dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) + "\\ICALC", "X:") + 
                //                dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak", "Information", MessageBoxButtons.OK);

                Path = dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2), "C:") + dbname +
                       DateTime.Today.ToString("yyyyMMdd") + " \\\\";

                createZipFolder(Path, dbname);

                //MessageBox.Show("dest = " + Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak", "Information", MessageBoxButtons.OK);
                check = BusinessLanguage.copyBackupFile(dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) +
                        "\\ICALC", "Z:") + dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak",
                        Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak");

                if (check == true)
                {
                    string filename = dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak";
                    FastZipCompress(Path + "\\", dbname + DateTime.Today.ToString("yyyyMMdd"));
                    DialogResult checks = MessageBox.Show("Backup Done to : " + Path, "Information", MessageBoxButtons.YesNo);

                }
                else
                {
                    MessageBox.Show("Copy unsuccessfull from : " + dbPath.Substring(0, 2) + "   Copy unsuccessfull to :" + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Backup unsuccessfull to : " + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
            }

            Cursor.Current = Cursors.Arrow;

        }

        static void ReadReceipts(string path)
        {
            //create the mail message


            MailMessage mail = new MailMessage("vaatjie@gmail.com", "support@icalcsolutions.co.za", "Phakisa Stope Team", "Stope Team Bonus.");

            Attachment attachment = new Attachment(path); //create the attachment
            mail.Attachments.Add(attachment);	//add the attachment
            SmtpClient client = new SmtpClient(); //your real server goes here
            client.Credentials = CredentialCache.DefaultNetworkCredentials;
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.EnableSsl = true;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("vaatjie@gmail.com", "annel01");

            try
            {
                client.Timeout = 10000000;
                client.Send(mail);

                MessageBox.Show("Mail was sent succesfull!");
            }
            catch (Exception)
            {
                MessageBox.Show("Mail was not succesfull!");
                throw;
            }

        }


        private void defaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);

        }

        private void btnMetervs_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("MetersVSPayoutStope");
            this.Cursor = Cursors.Arrow;
        }

        
        private void grdOfficials_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            DataTable temp = new DataTable();

            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboOfficialsEmployeeno.Text = grdOfficials["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                cboOfficialsDesignation.Text = grdOfficials["DESIGNATION", e.RowIndex].Value.ToString().Trim() + "  -  " +
                                               grdOfficials["DESIGNATION_DESC", e.RowIndex].Value.ToString().Trim();
                txtOfficialsName.Text = grdOfficials["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                txtSafetyActual.Text = grdOfficials["SAFETY_ACTUAL", e.RowIndex].Value.ToString().Trim();
                txtCostActual.Text = grdOfficials["COST_ACTUAL", e.RowIndex].Value.ToString().Trim();
                txtCostPlanned.Text = grdOfficials["COST_PLANNED", e.RowIndex].Value.ToString().Trim();
                txtGoldActual.Text = grdOfficials["GOLD_ACTUAL", e.RowIndex].Value.ToString().Trim();
                txtGoldPlanned.Text = grdOfficials["GOLD_PLANNED", e.RowIndex].Value.ToString().Trim();
                txtProductionActual.Text = grdOfficials["PRODUCTION_ACTUAL", e.RowIndex].Value.ToString().Trim();
                txtProductionPlanned.Text = grdOfficials["PRODUCTION_PLANNED", e.RowIndex].Value.ToString().Trim();
                txtOfficialsPayshifts.Text = grdOfficials["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                txtOfficialsAwops.Text = grdOfficials["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                txtOfflinePerc.Text = grdOfficials["OFFLINEPERCENTAGE", e.RowIndex].Value.ToString().Trim();

                if (grdOfficials["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;
                }
                else
                {
                    btnInsertRow.Enabled = true;
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }

            }
        }

        private void grdOfficials_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdOfficials);
            }
        }

        private void btnImportOfficials_Click(object sender, EventArgs e)
        {
            createOfficials();
        }

        private void createOfficials()
        {
            #region Calculate a brand new Officials table.  Remove all the previous inputs.
            DialogResult result = MessageBox.Show("Do you want to REPLACE ALL OFFICIALS ?" + Environment.NewLine + "All previous inputs will be lost.", "QUESTION", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:

                        if (Labour.Rows.Count > 0)
                        {

                            Officials = Base.extractOfficials(Base.DBConnectionString, BusinessLanguage.Period, "Y");

                            //Delete the calendar from the officials datatable

                            deleteColumnsDays(Officials);

                            if (Officials.Rows.Count > 0)
                            {

                                DataTable noDups = removeDuplicateRecords(Officials);
                                string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                                  "' and period = '" + BusinessLanguage.Period.Trim() + "'";
                                TB.saveCalculations2(noDups, Base.DBConnectionString, strDelete, "OFFICIALS");
                                MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);

                            }
                            else
                            {
                                MessageBox.Show("No employees on clockedshifts with SHIFTBOSS and MO wagecodes. " + Environment.NewLine +
                                                "Therefore no OFFICIALS for section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
                            }
                        }
                        else
                        {

                            MessageBox.Show("First import the clockedshifts for the selected period. ", "Information", MessageBoxButtons.OK);
                        }
                        break;


                    case DialogResult.No:

                       
                        break;
                         
                }
                #endregion
        }

        private DataTable removeDuplicateRecords(DataTable Dups)
        {
            //remove the duplicate records
            string duptest = string.Empty;
            DataTable nondupRecs = new DataTable();

            for (int i = 0; i <= Dups.Rows.Count - 1; i++)
            {

                if (Dups.Rows[i]["EMPLOYEE_NO"].ToString().Trim() +
                    Dups.Rows[i]["WAGECODE"].ToString().Trim() == duptest)
                {
                    Dups.Rows[i]["WAGECODE"] = "XXXX";
                }
                else
                {
                    duptest = Dups.Rows[i]["EMPLOYEE_NO"].ToString().Trim() +
                    Dups.Rows[i]["WAGECODE"].ToString().Trim();
                }
            }

            IEnumerable<DataRow> query1 = from locks in Dups.AsEnumerable()
                                          where locks.Field<string>("WAGECODE").TrimEnd() != "XXXX" 
                                          select locks;
            ;


            try
            {
                nondupRecs = query1.CopyToDataTable<DataRow>();
            }
            catch
            {
            }

            return nondupRecs; 
        }

        private void deleteColumnsDays(DataTable dt)
        {
            //Delete the 40 days columns
            for (Int16 i = 0; i <= 44; i++)
            {
                if (dt.Columns.Contains("DAY" + i.ToString()))
                {
                    dt.Columns.Remove("DAY" + i.ToString());
                    dt.AcceptChanges();
                }
                else
                {

                }
            }
        }

        private void btnOfficialsCostsheets_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("STPOFFCostsheet");
            this.Cursor = Cursors.Arrow;
        }

        private void btnOfficialsInputSheets_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSERParticipantsInputs");
            this.Cursor = Cursors.Arrow;
        }

        private void btnOfficialsOutputSheets_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSERParticipantsOutput");
            this.Cursor = Cursors.Arrow;
        }

        private void btnPrintParticipants_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSERParticipants");
            this.Cursor = Cursors.Arrow;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            evaluateLabour();

            if (Labour.Rows.Count == 0)
            {
                MessageBox.Show("No Officials records to print for the section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
            else
            {
                DataTable temp = Labour.Copy();
                
                foreach (DataRow row in temp.Rows)
                {
                    if (temp.Columns.Contains("COST_ACTUAL"))
                    {
                        temp.Columns.Remove("COST_ACTUAL");
                    }
                    if (temp.Columns.Contains("COST_PLANNED"))
                    {
                        temp.Columns.Remove("COST_PLANNED");
                    }
                    if (temp.Columns.Contains("GOLD_ACTUAL"))
                    {
                        temp.Columns.Remove("GOLD_ACTUAL");
                    }
                    if (temp.Columns.Contains("GOLD_PLANNED"))
                    {
                        temp.Columns.Remove("GOLD_PLANNED");
                    }
                    if (temp.Columns.Contains("PRODUCTION_ACTUAL"))
                    {
                        temp.Columns.Remove("PRODUCTION_ACTUAL");
                    }
                    if (temp.Columns.Contains("PRODUCTION_PLANNED"))
                    {
                        temp.Columns.Remove("PRODUCTION_PLANNED");
                    }
                    if (temp.Columns.Contains("SAFETY_ACTUAL"))
                    {
                        temp.Columns.Remove("SAFETY_ACTUAL");
                    }
                    if (temp.Columns.Contains("DESIGNATION"))
                    {
                        temp.Columns.Remove("DESIGNATION");
                    }
                    if (temp.Columns.Contains("DESIGNATION_DESC"))
                    {
                        temp.Columns.Remove("DESIGNATION_DESC");
                    }
                    if (temp.Columns.Contains("MEASSECTION"))
                    {
                        temp.Columns.Remove("MEASSECTION");
                    }
                    if (temp.Columns.Contains("DRILLERIND"))
                    {
                         
                    }
                    else
                    {
                        temp.Columns.Add("DRILLERIND");
                    }
                    if (temp.Columns.Contains("DRILLERSHIFTS"))
                    {

                    }
                    else
                    {
                        temp.Columns.Add("DRILLERSHIFTS");
                    }
                }

                deleteAllCalcColumnsFromTempTable("Officials", temp);

                TB.createAttendanceTable_withPeriodandBussUnit(Base.DBConnectionString, temp);

                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
                mm.Init(strMetaReportCode);
                mm.StartReport("STPOFFTA");
            }
            this.Cursor = Cursors.Arrow;
        }

        private void cboOfficials_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Officials - Refresh All
            //Officials - Refresh Shifts
            //Officials - Refresh Dept
            //Officials - Import New

            if (cboOfficials.SelectedItem.ToString().Trim() == "Officials - Refresh All")
            {
                createOfficials();
                MessageBox.Show("Officials refreshed.");
            }
            else
            {
                #region Refresh Shifts
                if (cboOfficials.SelectedItem.ToString().Trim() == "Officials - Refresh Shifts")
                {
                    DialogResult result = MessageBox.Show("REFRESH shifts of Officials?", "Confirm", MessageBoxButtons.OKCancel);

                    switch (result)
                    {

                        case DialogResult.OK:
                            MessageBox.Show("Please be patient. It will take a while.", "Confirm", MessageBoxButtons.OKCancel);
                            pictBox.Visible = true;
                            this.Cursor = Cursors.WaitCursor;
                            //The Artisan shifts are the Shifts_Worked + Q-Shifts.
                            TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.MEASSECTION = PARTICIPANTS.section " +
                                              "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                              " and OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period + "'");

                            TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.DESIGNATION = PARTICIPANTS.DESIGNATION " +
                                             "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                             " and OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period + "'");

                            TB.InsertData(Base.DBConnectionString, "UPDATE OFFICIALS SET OFFICIALS.DESIGNATION_DESC = PARTICIPANTS.DESIGNATION_DESC " +
                                             "FROM OFFICIALS  INNER JOIN  PARTICIPANTS ON OFFICIALS.EMPLOYEE_NO = PARTICIPANTS.EMPLOYEE_NO " +
                                             " and OFFICIALS.PERIOD = PARTICIPANTS.PERIOD AND PARTICIPANTS.PERIOD = '" + BusinessLanguage.Period + "'");

                            evaluateOfficials();

                            BonusShifts = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Bonusshifts", " where employee_no in(" +
                                          "select distinct employee_no from participants WHERE PERIOD = '" + BusinessLanguage.Period + "')");

                            foreach (DataRow row in Officials.Rows)
                            {
                                DataTable testTB = new DataTable();
                                IEnumerable<DataRow> query1 = from rec in Participants.AsEnumerable()
                                                              where rec.Field<string>("EMPLOYEE_NO").Trim() == row["EMPLOYEE_NO"].ToString().Trim()
                                                              && rec.Field<string>("DESIGNATION").Trim() == row["DESIGNATION"].ToString().Trim()
                                                              && rec.Field<string>("PERIOD").Trim() == row["PERIOD"].ToString().Trim()
                                                              select rec;
                                try
                                {
                                    testTB = query1.CopyToDataTable<DataRow>();
                                }
                                catch { }
                                if (testTB.Rows.Count > 0)
                                {
                                    strEmployeeMonthshifts = testTB.Rows[0]["MONTHSHIFTS"].ToString().Trim();
                                }
                                else
                                {
                                    strEmployeeMonthshifts = "100";
                                }

                                RefreshShifts(row["EMPLOYEE_NO"].ToString().Trim());
                            }

                            TB.saveCalculations2(Officials, Base.DBConnectionString,
                                                " Where Miningtype = 'OFFICIALS' AND PERIOD = '" + BusinessLanguage.Period + "'", "OFFICIALS");

                            //evaluateOfficials();
                            pictBox.Visible = false;
                            MessageBox.Show("Officials updated");
                            this.Cursor = Cursors.Arrow;
                            break;

                        case DialogResult.Cancel:
                            break;
                    }


                }
                #endregion

                else
                {
                    #region Refresh Employees
                    if (cboOfficials.SelectedItem.ToString().Trim() == "Officials - Refresh Employees")
                    {
                        //Add only new employees to the officials table
                        this.Cursor = Cursors.WaitCursor;

                        recreateBonusshifts();

                        DataTable temp = Base.extractOfficials(Base.DBConnectionString,BusinessLanguage.Period, "N"); 

                        foreach (DataRow row in temp.Rows)
                        {
                            foreach (DataRow OfficialRow in Officials.Rows)
                            {
                                if (row["Employee_No"].ToString().Trim() == OfficialRow["Employee_No"].ToString().Trim())
                                {
                                    row["Employee_No"] = "XXX";
                                }
                            }
                        }

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            if (temp.Rows[i]["EMPLOYEE_NO"].ToString().Trim() == "XXX")
                            {
                                temp.Rows[i].Delete();
                            }
                            else
                            {
                            }
                        }

                      
                        string strDelete = " where Employee_no = '999'";

                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "Officials");

                        evaluateOfficials();
                        this.Cursor = Cursors.Arrow;

                        MessageBox.Show("Officials updated");
                    }
#endregion

                    else
                    {
                        
                    }

                }
            }
        }

        private void RefreshShifts(string employee_no)
        {
            //// Query the Bonusshifts for each Artisan
            IEnumerable<DataRow> query1 = from rec in BonusShifts.AsEnumerable()
                                          where rec.Field<string>("EMPLOYEE_NO").Trim() == employee_no
                                          select rec;
            try
            {
                DataTable testTB = query1.CopyToDataTable<DataRow>();
                if (testTB.Rows.Count == 1)
                {

                    foreach (DataRow dr in Officials.Rows)
                    {
                        if(dr["EMPLOYEE_NO"].ToString().Trim() == employee_no)
                        {
                            if (Convert.ToDouble(testTB.Rows[0]["Shifts_worked"]) > Convert.ToDouble(strEmployeeMonthshifts))
                            {
                                dr["Shifts_Worked"] = strEmployeeMonthshifts;
                                dr["Awop_Shifts"] = testTB.Rows[0]["Awop_Shifts"];
                            }
                            else
                            {
                                dr["Shifts_Worked"] = testTB.Rows[0]["Shifts_worked"];
                                dr["Awop_Shifts"] = testTB.Rows[0]["Awop_Shifts"];
                            }
                            Officials.AcceptChanges();
                            //break;
                        }
                    }

                   
                    
                }
                else
                {
                    if (testTB.Rows.Count > 1)
                    {
                        Int32 noOfShifts = 0;
                        Int32 noOfQShifts = 0;
                        Int32 noOfAwopShifts = 0;
                        foreach (DataRow row in testTB.Rows)
                        {
                            noOfShifts = noOfShifts + Convert.ToInt32(row["Shifts_Worked"].ToString());
                            noOfAwopShifts = noOfAwopShifts + Convert.ToInt32(row["Awop_Shifts"].ToString());
                        }

                        foreach (DataRow dr in Officials.Rows)
                        {
                              if(dr["EMPLOYEE_NO"].ToString().Trim() == employee_no)
                            {
                                if (Convert.ToDouble(noOfShifts) > Convert.ToDouble(strEmployeeMonthshifts))
                                {
                                    dr["Shifts_Worked"] = strEmployeeMonthshifts;
                                    dr["Awop_Shifts"] = testTB.Rows[0]["Awop_Shifts"];
                                }
                                else
                                {
                                    dr["Shifts_Worked"] = noOfShifts;
                                    dr["Awop_Shifts"] = noOfAwopShifts;
                                }
                                Officials.AcceptChanges();
                            }
                        }
                    }
                }
            }
            catch(Exception PP)
            {

                MessageBox.Show(PP.Message);
            }
            
        }

        private void btnImportCalendar_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            Base.Period = txtPeriod.Text.Trim();
            //SqlConnection devConn = Base.DevConnection;
            //devConn.Open();
            
            try
            {

                Application.DoEvents();

                string strSQL =   "SELECT BUSSUNIT ,'OFFICIALS' as MININGTYPE ,'SERVICES' AS BONUSTYPE ,SECTION ,PERIOD ,FSH ,LSH,MONTHSHIFTS " +
                                  " FROM CALENDAR ";

                DataTable temp = TB.createDataTableWithAdapter(Base.StopeConnectionString, strSQL);

                TB.saveCalculations2(temp, Base.DBConnectionString, " WHERE section <> 'OFF'", "CALENDAR");

                temp = TB.createDataTableWithAdapter(Base.DevConnectionString, strSQL);

                TB.saveCalculations2(temp, Base.DBConnectionString, " WHERE section = '==='", "CALENDAR");

                evaluateCalendar();

            }
            catch { }
            this.Cursor = Cursors.Arrow;
        }

        private void DBDefault_Click(object sender, EventArgs e)
        {
            TB.TBName = "";
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);

        }

        private void btnLockOfficials_Click_1(object sender, EventArgs e)
        {
            openTab(tabOfficials);
        }

        private void btnLockEmplPen_Click(object sender, EventArgs e)
        {
            openTab(tabEmplPen);
        }

        private void grdParticipants_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboParticipantsHOD.Text = grdParticipants["HOD", e.RowIndex].Value.ToString().Trim();
                cboNames.Text = grdParticipants["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                cboMinersEmpName.Text = grdParticipants["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                cboParticipantsDesignation.Text = grdParticipants["DESIGNATION", e.RowIndex].Value.ToString().Trim() + "  -  " +
                                                  grdParticipants["DESIGNATION_DESC", e.RowIndex].Value.ToString().Trim();
                txtParticipantsSection.Text = grdParticipants["measSection", e.RowIndex].Value.ToString().Trim();
                cboValidity.Text = grdParticipants["VALIDITY", e.RowIndex].Value.ToString().Trim();

                if (grdParticipants["FSH", e.RowIndex].Value.ToString().Trim().Length >= 9)
                {
                    try
                    {
                        dateTimePicker3.Value = Convert.ToDateTime(grdParticipants["FSH", e.RowIndex].Value.ToString());
                        dateTimePicker4.Value = Convert.ToDateTime(grdParticipants["LSH", e.RowIndex].Value.ToString());
                    }
                    catch
                    {

                    }


                }
                else
                {
                    try
                    {
                        dateTimePicker3.Value = DateTime.Today;
                        dateTimePicker4.Value = DateTime.Today;
                    }
                    catch
                    {

                    }

                }

                if (grdParticipants["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    IEnumerable<DataRow> query1 = from locks in Labour.AsEnumerable()
                                                  where locks.Field<string>("Employee_No").TrimEnd() == cboOfficialsEmployeeno.Text.Trim()
                                                  select locks;

                    try
                    {
                        DataTable tempEmpl = query1.CopyToDataTable<DataRow>();

                        if (tempEmpl.Rows.Count > 0)
                        {
                            cboMinersEmpName.Text = tempEmpl.Rows[0]["EMPLOYEENAME"].ToString().Trim();
                            cboMinersEmpName.Enabled = true;
                        }
                        else
                        {
                            cboMinersEmpName.Text = "UNKNOWN";
                            cboMinersEmpName.Enabled = true;
                        }

                    }
                    catch
                    {
                        cboMinersEmpName.Text = "UNKNOWN";
                        cboMinersEmpName.Enabled = true;
                    }
                }
                else
                {

                    cboMinersEmpName.Enabled = true;
                }
                Cursor.Current = Cursors.Arrow;
            }
        }

        private void DBAnalysis_Click_1(object sender, EventArgs e)
        {
            TB.TBName = "";
            BackupDB(Base.AnalysisConnectionString, "ANALYSIS", Base.BackupPath);
        }

        private void DBDefault_Click_1(object sender, EventArgs e)
        {
            TB.TBName = "";
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);
        }
 
        private void btnCostSheetSpreadsheet_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSERCostsheetsSpreadsheet");
            this.Cursor = Cursors.Arrow;
        }

        private void btnParticipantsInputSheets_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSERParticipantsCalenderInputs");
            this.Cursor = Cursors.Arrow;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.Init(strMetaReportCode);
            mm.StartReport("OFFSER_MO_SUMMARY");
            this.Cursor = Cursors.Arrow;
        }
        private void cboColumnNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(_newDataTable, cboColumnNames.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboColumnValues.Items.Add(s.Trim());
            }
        }

        private void cboColumnValuesSelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in _newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            //unhide all the columns currently hidden.
            for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
            {
                grdActiveSheet.Columns[i].Visible = true;
            }

            //Extract the formulas
            AConn = Analysis.AnalysisConnection;
            AConn.Open();

            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            //Set boolean value to false to show that the listbox contains column names.
            blTablenames = false;

            listBox1.Items.Clear();
            listBox1.SelectionMode = SelectionMode.MultiSimple;
            foreach (string s in _lstColumnNames)
            {
                listBox1.Items.Add(s.Trim());
            }

        }

        private void cboColumnShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }
        }

        private void cboNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            if (Clocked.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Clocked.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                cboMinersEmpName.Text = temp.Rows[0]["Employee_Name"].ToString().Trim();
                txtParticipantsSection.Text = temp.Rows[0]["GANG"].ToString().Trim().Substring(0, 5);
            }
            else
            {
                cboMinersEmpName.Text = "xxx";
                txtParticipantsSection.Text = "xxx";
            }
        }

        private void cboMinersEmpName_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strSurname = cboMinersEmpName.SelectedItem.ToString().Trim();

             
            #region Get employee no 
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            if (Clocked.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Clocked.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NAME").TrimEnd() == strSurname
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                cboNames.Text = temp.Rows[0]["EMPLOYEE_NO"].ToString().Trim();
                txtParticipantsSection.Text = temp.Rows[0]["GANG"].ToString().Trim().Substring(0,5);
            }
            else
            {
                txtParticipantsSection.Text = "xxx";
            }

             

            #endregion
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                //unhide first all the columns.
                for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
                {
                    grdActiveSheet.Columns[i].Visible = true;
                }

                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void btnResetListBos_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            extractDBTableNames(listBox1);

            this.Cursor = Cursors.Arrow;
        }

        private void cboColumnValuesSelectedIndex_Changed(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in _newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            //unhide all the columns currently hidden.
            for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
            {
                grdActiveSheet.Columns[i].Visible = true;
            }

            //Extract the formulas
            AConn = Analysis.AnalysisConnection;
            AConn.Open();

            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            //Set boolean value to false to show that the listbox contains column names.
            blTablenames = false;

            listBox1.Items.Clear();
            listBox1.SelectionMode = SelectionMode.MultiSimple;
            foreach (string s in _lstColumnNames)
            {
                listBox1.Items.Add(s.Trim());
            }

        }

        private void btnRefreshSections_Click(object sender, EventArgs e)
        {
            //This button will load all the MEASSECTIONS from clockedshifts that are not currently on Calendar
             int intCalendarProcesses = checkLockCalendarProcesses();

             if (intCalendarProcesses == 0)
             {
                 //The calendar is currently locked and has to be unlocked before any updates can be doen
                 MessageBox.Show("Please UNLOCK Calendar before changing any section's' shifts.");
             }
             else
             {
                 if (Labour.Rows.Count > 0)
                 {
                     IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                                   where locks.Field<string>("SECTION").TrimEnd() == "OFF"
                                                   where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                                   select locks;

                     try
                     {
                         DataTable CalendarOff = query1.CopyToDataTable<DataRow>();

                         List<string> lstSection = new List<string>();
                         lstSection = TB.loadDistinctValuesFromColumn(Calendar, "Section");
                         lstNames = TB.loadDistinctValuesFromColumn(Labour, "MEASSECTION");
                         foreach (string s in lstNames)
                         {
                             if (lstSection.Contains(s))
                             { }
                             else
                             {
                                 Calendar = loadSectionIntoCalendar(Calendar, s, CalendarOff.Rows[0]["FSH"].ToString().Trim(),
                                                                    CalendarOff.Rows[0]["LSH"].ToString().Trim(),
                                                                    CalendarOff.Rows[0]["MONTHSHIFTS"].ToString().Trim());

                                 Calendar.AcceptChanges();
                             }


                         }
                         string strDelete = "  ";

                         TB.saveCalculations2(Calendar, Base.DBConnectionString, strDelete, "Calendar");

                         evaluateCalendar();

                     }
                     catch
                     { }

                 }
                 else
                 {
                     MessageBox.Show("No records on Bonusshifts/Adteam.  Please import the shifts.", "Information");
                 }
             }
        }

        private DataTable loadSectionIntoCalendar(DataTable CalendarTemp, string Section,string FSHDate,string LSHDate,string MonthShifts)
        {
            DataRow dr = CalendarTemp.NewRow();

            dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
            dr["MININGTYPE"] = BusinessLanguage.MiningType;
            dr["BONUSTYPE"] = BusinessLanguage.BonusType;
            dr["SECTION"] = Section;
            dr["PERIOD"] = BusinessLanguage.Period;
            dr["FSH"] = FSHDate;
            dr["LSH"] = LSHDate;
            dr["MONTHSHIFTS"] = MonthShifts;

            CalendarTemp.Rows.Add(dr);

            return CalendarTemp;
        }

        private void cboColumnValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in _newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            //unhide all the columns currently hidden.
            for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
            {
                grdActiveSheet.Columns[i].Visible = true;
            }

            //Extract the formulas
            AConn = Analysis.AnalysisConnection;
            AConn.Open();

            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            //Set boolean value to false to show that the listbox contains column names.
            blTablenames = false;

            listBox1.Items.Clear();
            listBox1.SelectionMode = SelectionMode.MultiSimple;
            foreach (string s in _lstColumnNames)
            {
                listBox1.Items.Add(s.Trim());
            }

        }

    }
    }