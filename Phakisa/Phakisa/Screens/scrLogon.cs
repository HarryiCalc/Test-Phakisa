using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using TB = clsTable;
using Base = clsMain;
namespace Phakisa
{
    public partial class scrLogon : Form
    {
        InputBox InputBox = new InputBox();
        clsBL.clsBL bl = new clsBL.clsBL();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsDBase.clsDBase DB = new clsDBase.clsDBase();
        clsMain.clsMain Base = new clsMain.clsMain();
        SqlConnection BaseConn = new SqlConnection();
        SqlConnection EnvConn = new SqlConnection();
        DataTable BussMinBon = new DataTable();
        DataTable BonusTypes = new DataTable();
        DataTable User = new DataTable();
        string strBonusModule = "";
        string Drive = string.Empty;
        string strServerPath = string.Empty;

        public scrLogon()
        {
            InitializeComponent();
        }
        private void scrLogon_Load(object sender, EventArgs e)
        {
            panel1.Enabled = false;
            panel4.Enabled = false;
            panelSignon.Enabled = false;
            panel3.Enabled = false;
            label10.Visible = false;

            cboUserid.Items.Clear();
            cboBussUnit.Items.Clear();
            cboBussUnit.Focus();



            //Encoding.Unicode.GetString(Convert.FromBase64String(Base.AnyWhere)).Trim())
            //+ Encoding.Unicode.GetString(Convert.FromBase64String(Base.Setinit)).Trim()

            if (Base.checkIfFileExists(Encoding.Unicode.GetString(Convert.FromBase64String(Base.AnyWhere)).Trim() +
                "\\" + Encoding.Unicode.GetString(Convert.FromBase64String(Base.Setinit)).Trim()))
            {
            }
            else
            {
                MessageBox.Show("This is an illigal copy of the system.", "Warning", MessageBoxButtons.OK);
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool checkId = true;
            bool checkPs = true;
            int dateDiff = 0;

            if (cboUserid.Text.Length > 0 && cboBussUnit.Text.Length > 0 && txtRegion.Text.Length > 0 && txtPassw.Text.Length > 0)
            {
                User = Base.createDataTableWithAdapterSelectAll("Profile", "where userid = '" + cboUserid.Text.Trim() + "' and environment = '" + cboEnvironment.Text.Trim() + "'", Base.BaseConnectionString);

                //object intCount = Base.CountUserid(cboUserid.SelectedItem.ToString().Trim().ToUpper(), BaseConn);

                if (User.Rows.Count > 0)
                {
                    //Userid was found.

                    object intcount2 = Base.CountUseridPassword(cboUserid.Text.ToString().Trim().ToUpper(), txtPassw.Text.Trim(), cboEnvironment.Text.Trim(), Base.BaseConnectionString);//Take out for live program
                    //Put back after testing for life program
                    intcount2 = Base.CountUseridPassword(cboUserid.SelectedItem.ToString().Trim().ToUpper(), txtPassw.Text.Trim(), cboEnvironment.Text.Trim(), Base.BaseConnectionString);

                    if ((int)intcount2 > 0)
                    {

                    }
                    else
                    {
                        txtPassw.Text = "";
                        MessageBox.Show("Invalid password", "Logon Failed", MessageBoxButtons.OK);
                        checkPs = false;
                    }
                }
                else
                {
                    MessageBox.Show("Invalid userid", "Logon Failed", MessageBoxButtons.OK);
                    checkId = false;
                }

                if (checkId == true && checkPs == true)
                {
                    bl.Userid = cboUserid.Text.ToString().Trim().ToUpper();
                    bl.Region = txtRegion.Text.ToString().Trim().ToUpper();
                    bl.BussUnit = cboBussUnit.Text.ToString().Trim().Substring(0, cboBussUnit.Text.ToString().Trim().IndexOf(" - "));

                    if (bl.Userid == "ADMIN")
                    {
                        Base.UpdateADMIN(bl.BussUnit, Base.BaseConnectionString);
                    }

                    DateTime today = DateTime.Today;

                    object dtePasswordDte = Base.selectPasswordDate(bl.Userid, txtPassw.Text.Trim(), cboEnvironment.Text.Trim(), Base.BaseConnectionString);

                    DateTime passwordDate = Convert.ToDateTime((string)dtePasswordDte);

                    if ((today.Year - passwordDate.Year) < 0)
                    {
                        dateDiff = -10;

                    }
                    else
                    {
                        if ((today.Year - passwordDate.Year) > 0)
                        {
                            dateDiff = 10;

                        }
                        else
                        {
                            dateDiff = today.DayOfYear - passwordDate.DayOfYear;
                        }
                    }

                    if (dateDiff > 0)
                    {
                        InputBoxResult result = InputBox.Show("Password has expired. Enter a new password: ");

                        if (result.ReturnCode == DialogResult.OK)
                            if (result.Text.Trim() == "")
                            {
                            }
                            else
                            {
                                string check = bl.validatePassword(result.Text.Trim());

                                if (check.Length > 0)
                                {
                                    MessageBox.Show(check.Trim(), "Validation Message", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    DateTime dteTemp = today.AddDays(30);
                                    Base.UpdatePassword(result.Text.Trim(), dteTemp.ToString("yyyy-MM-dd"), bl.Userid.Trim(), Base.BaseConnectionString);
                                    panel1.Enabled = true;
                                    BussMinBon = Base.getBussMinBon(cboBussUnit.SelectedItem.ToString().Trim().Substring(0, cboBussUnit.SelectedItem.ToString().Trim().IndexOf(" - ")), txtRegion.Text.Trim(), Base.BaseConnectionString);
                                    loadCombos(BussMinBon);
                                }
                            }
                        else
                        {
                        }
                    }
                    else
                    {
                        //Hier kom die activering.....
                        BussMinBon = Base.getBussMinBon(cboBussUnit.Text.ToString().Trim().Substring(0, cboBussUnit.Text.ToString().Trim().IndexOf(" - ")), txtRegion.Text.Trim(), Base.BaseConnectionString);
                        loadCombos(BussMinBon);
                        //Disable all the dropboxes
                        txtRegion.Enabled = false;
                        cboBussUnit.Enabled = false;
                        cboUserid.Enabled = false;
                        txtPassw.Enabled = false;
                        panelSignon.Enabled = false;
                        cboBonusType.Enabled = false;
                        cboPeriods.Enabled = false;
                        cboMiningType.Focus();
                        //this.Close();

                    }
                }
            }
            else
            {

                MessageBox.Show("Please supply all input necessary.", "Signon Information.", MessageBoxButtons.OK);

            }
        }

        private void loadCombos(DataTable BussMinBon)
        {
            panel1.Enabled = true;
            List<string> lstNames = TB.loadDistinctValuesFromColumn(BussMinBon, "MININGTYPE");
            cboMiningType.Items.Clear();
            foreach (string s in lstNames)
            {
                cboMiningType.Items.Add(s.Trim());
            }

            //cboMiningType.Items.Add("DEVELOPMENT");
            cboMiningType.Text = cboMiningType.Items[0].ToString();
            cboMiningType.BackColor = Color.Cornsilk;
            cboBonusType.BackColor = Color.Cornsilk;
            cboPeriods.BackColor = Color.Cornsilk;
            //cboMiningType.Text = "DEVELOPMENT";
            cboMiningType_SelectedIndexChanged("Method", null);


        }

        private void Execute()
        {
            switch (cboMiningType.Text.Trim())
            {
                case "OFFICIALS":
                    scrOfficials officials = new scrOfficials();
                    officials.scrOfficialsLoad(cboPeriods.Text.Trim(), txtRegion.Text.Trim(), (string)cboBussUnit.Text.Substring(0, cboBussUnit.Text.Trim().IndexOf("-")).Trim(), cboUserid.Text.Trim(), cboMiningType.Text.Trim(), cboBonusType.Text.Trim(), cboEnvironment.Text.Trim());
                    this.Hide();
                    officials.Show();
                    break;

                case "DEVELOPMENT":
                    scrTeamD teams = new scrTeamD();
                    teams.scrTeamDLoad(cboPeriods.Text.Trim(), txtRegion.Text.Trim(), (string)cboBussUnit.Text.Substring(0, cboBussUnit.Text.Trim().IndexOf("-")).Trim(), cboUserid.Text.Trim(), cboMiningType.Text.Trim(), cboBonusType.Text.Trim(), cboEnvironment.Text.Trim());
                    this.Hide();
                    teams.Show();
                    break;

                case "ENGINEERING":
                    scrEngineering Engineering = new scrEngineering();
                    Engineering.scrEngineeringLoad(cboPeriods.Text.Trim(), txtRegion.Text.Trim(), (string)cboBussUnit.Text.Substring(0, cboBussUnit.Text.Trim().IndexOf("-")).Trim(), cboUserid.Text.Trim(), cboMiningType.Text.Trim(), cboBonusType.Text.Trim(), cboEnvironment.Text.Trim());
                    this.Hide();
                    Engineering.Show();
                    break;


                case "MPB":
                    MessageBox.Show("Please choose 'Stope' as your mining type.");
                    break;

                case "TRANSPORT":
                    MessageBox.Show("Please choose 'Stope' as your mining type.");
                    break;
            }
        }

        private void cboBussUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Extract all the userids for the selected REGION and BUSSUNIT

            DataTable Userids = Base.getUseridByBussUnit(cboBussUnit.SelectedItem.ToString().Trim().Substring(0, cboBussUnit.SelectedItem.ToString().Trim().IndexOf(" - ")), Base.BaseConnectionString);

            cboUserid.Items.Clear();

            List<string> distinctUserids = TB.loadDistinctValuesFromColumn(Userids, "Userid");

            foreach (string s in distinctUserids)
            {
                cboUserid.Items.Add(s.ToString().Trim());
            }

            cboUserid.Enabled = true;
            cboPeriods.Enabled = true;

        }

        private void cboUserid_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtPassw.Enabled = true;

            txtPassw.Focus();

        }

        private void cboMiningType_SelectedIndexChanged(object sender, EventArgs e)
        {
            BonusTypes = Base.createDataTableWithAdapterSelectAll("BussMinBon", "where miningtype = '" + cboMiningType.Text.Trim() +
                                                                    "' and bussunit = '" +
                                                                    cboBussUnit.SelectedItem.ToString().Trim().Substring(0, cboBussUnit.SelectedItem.ToString().Trim().IndexOf(" - ")).Trim() + "'",
                                                                    Base.BaseConnectionString);
            cboBonusType.Items.Clear();
            foreach (DataRow r in BonusTypes.Rows)
            {
                if (cboBonusType.Items.Contains(r["BONUSTYPE"].ToString().Trim()))
                {

                }
                else
                {
                    cboBonusType.Items.Add(r["BONUSTYPE"].ToString().Trim());
                }

            }
            cboBonusType.Enabled = true;
            cboBonusType.Focus();

        }

        private void cboBonusType_SelectedIndexChanged(object sender, EventArgs e)
        {

            IEnumerable<DataRow> query1 = from bonusmodules in BussMinBon.AsEnumerable()
                                          where bonusmodules.Field<string>("MININGTYPE").TrimEnd() == cboMiningType.Text.Trim()
                                          where bonusmodules.Field<string>("BONUSTYPE").TrimEnd() == cboBonusType.Text.Trim()
                                          select bonusmodules;

            try
            {
                DataTable BonusModule = query1.CopyToDataTable<DataRow>();
                strBonusModule = BonusModule.Rows[0]["BONUSMODULE"].ToString();
                cboPeriods.Enabled = true;
                cboPeriods.Focus();
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "Error", MessageBoxButtons.OK);

            }
            //int intcount = query1.Count<DataRow>();

            //return intcount;


        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (cboMiningType.Text.Length > 0 && cboBonusType.Text.Length > 0 && cboPeriods.Text.Length > 0)
            {

                Execute();
            }
            else
            {

                MessageBox.Show("Please supply all input necessary.", "Signon Information.", MessageBoxButtons.OK);
            }

        }

        private void doMainDBExtracts()
        {

            //Extract all the regions
            DataTable bussunits = Base.SelectBussUnit(Base.BaseConnectionString);
            //List<string> distinctRegions = TB.loadDistinctValuesFromColumn(bussunits, "Region");

            //foreach (string s in distinctRegions)
            //{
            //    cboRegion.Items.Add(s.ToString().Trim());
            //}

            //Calculate the periods
            DateTime today = DateTime.Today;
            string strYear = Convert.ToString(today.Year - 1);

            for (int i = 8; i < 10; i++)
            {
                cboPeriods.Items.Add(strYear + "0" + Convert.ToString(i));
            }
            for (int i = 10; i < 13; i++)
            {
                cboPeriods.Items.Add(strYear + Convert.ToString(i));
            }

            strYear = today.Year.ToString();

            for (int i = 1; i < 10; i++)
            {
                cboPeriods.Items.Add(strYear + "0" + Convert.ToString(i));
            }
            for (int i = 10; i < 13; i++)
            {
                cboPeriods.Items.Add(strYear + Convert.ToString(i));
            }
            string strMonth = Convert.ToString(today.Month);

            if (strMonth.Trim().Length == 1)
            {

                cboPeriods.SelectedItem = strYear + "0" + strMonth.Trim();

            }

            else
            {

                cboPeriods.SelectedItem = strYear + strMonth.Trim();

            }
        }

        private void cboPeriods_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnProcess.Enabled = true;
            panel3.Enabled = true;
        }

        private void cboEnvironment_SelectedIndexChanged(object sender, EventArgs e)
        {
            label10.Text = "Environment = ";
            label10.Text = label10.Text + "  " + cboEnvironment.Text.Trim();
            label10.Visible = true;
            panel1.Enabled = true;
            panel4.Enabled = true;
            panelSignon.Enabled = true;

            setEnvironment();

            doMainDBExtracts();
            cboBussUnit.Items.Clear();
            cboUserid.Items.Clear();

            cboUserid.Enabled = false;
            cboBussUnit.Enabled = false;

            //Extract all the bussunits for the selected region
            //USE TO BE cboRegion....
            DataTable bussunits = Base.GetBussUnitByRegion(txtRegion.Text.Trim(), Base.BaseConnectionString);

            foreach (DataRow r in bussunits.Rows)
            {
                cboBussUnit.Items.Add(r["BussUnit"].ToString().Trim() + " - " + r["BussUnit_Desc"].ToString().Trim());
            }

            cboBussUnit.Enabled = true;


        }

        private void setEnvironment()
        {
            //Haal uit 
            btnProcess.Enabled = true;

            Base.Drive = System.Configuration.ConfigurationSettings.AppSettings[cboEnvironment.Text.Trim() + "Drive"];
            Base.Integrity = System.Configuration.ConfigurationSettings.AppSettings[cboEnvironment.Text.Trim() + "Integrity"];
            Base.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[cboEnvironment.Text.Trim() + "Userid"])).Trim();
            Base.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[cboEnvironment.Text.Trim() + "Password"])).Trim();
            Base.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[cboEnvironment.Text.Trim() + "ServerName"])).Trim();

            Base.BaseConnectionString = Base.ServerName.Trim();
            cboEnvironment.Enabled = false;
            cboBussUnit.Focus();

            if (cboEnvironment.Text.Trim() == "Development")
            {
                
                Base.BaseConnectionString = Environment.MachineName.Trim() + Base.ServerName;
                Base.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
            }

            // MessageBox.Show(Base.BaseConnectionString, "Information", MessageBoxButtons.OK);

            //MessageBox.Show(Share.BaseConnectionString);

        }

        private void scrLogon_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }



    }
}
