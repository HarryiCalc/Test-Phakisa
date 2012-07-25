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
using System.Collections;
using System.Resources;
using System.Data.OleDb;


namespace Phakisa
{
    public partial class scrPayroll : Form
    {
        string column = string.Empty;
        string globalSection = "";
        clsTableFormulas TBFormulas = new clsTableFormulas();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsMain.clsMain Base = new clsMain.clsMain();
        SqlConnection tstConn = new SqlConnection();
        SqlConnection BaseConn = new SqlConnection();
        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        DataTable dt = new DataTable();
        DataTable outputTable = new DataTable();
        DataTable earnCode = new DataTable();
        string strDay = "00";
        string strMonth = "00";
        List<string> lstNames = new List<string>();

        DataTable manualPay = new DataTable();
        DataTable newPayroll = new DataTable();

        CheckBox chkB;

        public scrPayroll()
        {
            InitializeComponent();
        }

        internal void PayrollSendLoad(SqlConnection myConn, SqlConnection BConn, clsBL.clsBL classBL, clsTable.clsTable classTable, clsTableFormulas classTableFormulas, clsMain.clsMain Main, string section)
        {
            if (section == "MANUAL")
            {

                globalSection = section;
                Base = Main;
                TBFormulas = classTableFormulas;
                TB = classTable;
                BusinessLanguage = classBL;

                tstConn = myConn;
                BaseConn = BConn;

                txtBussUnit.Text = BusinessLanguage.BussUnit + " - " + BusinessLanguage.Region;
                txtBussUnit.Enabled = false;
                cboPayrollGroup.Text = TB.TBName.Trim();

                DateTime today = DateTime.Today;
                strDay = strDay.Substring(0, 2 - today.Day.ToString().Trim().Length) + today.Day.ToString().Trim();
                strMonth = strMonth.Substring(0, 2 - today.Month.ToString().Trim().Length) + today.Month.ToString().Trim();


                cboMiningType.Text = BusinessLanguage.MiningType;

                cboBonusType.Text = BusinessLanguage.BonusType;  //amp
                cboSection.Text = globalSection;
                earnCode = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EARNINGSCODES", " Where Tablename = 'MANUALPAY'"); //amp
                List<string> employeeType = TB.loadDistinctValuesFromColumn(earnCode, "EMPLOYEETYPE");

                foreach (string s in employeeType)
                {
                    cboEmployeeType.Items.Add(s.ToString().Trim());
                }

                txtPeriod.Text = Base.Period;
                cboPayRollInd.Enabled = false;
                cboEarningsCode.Enabled = false;
                cboMiningType.Enabled = false;
                cboBonusType.Enabled = false;
                cboSection.Enabled = false;
                btnGenerate.Visible = false;

            }
            else
            {
                btnManual.Visible = false;
                globalSection = section;
                Base = Main;
                TBFormulas = classTableFormulas;
                TB = classTable;
                BusinessLanguage = classBL;

                tstConn = myConn;
                BaseConn = BConn;

                txtBussUnit.Text = BusinessLanguage.BussUnit + " - " + BusinessLanguage.Region;
                txtBussUnit.Enabled = false;
                cboPayrollGroup.Text = TB.TBName.Trim();

                DateTime today = DateTime.Today;
                strDay = strDay.Substring(0, 2 - today.Day.ToString().Trim().Length) + today.Day.ToString().Trim();
                strMonth = strMonth.Substring(0, 2 - today.Month.ToString().Trim().Length) + today.Month.ToString().Trim();


                cboMiningType.Text = BusinessLanguage.MiningType;

                cboBonusType.Text = BusinessLanguage.BonusType;

                cboSection.Text = globalSection;


                //dt = Base.GetDataByRegion(BusinessLanguage.Region, Base.DBConnectionString);
                earnCode = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EARNINGSCODES", "Where Tablename != 'MANUALPAY'");
                List<string> employeeType = TB.loadDistinctValuesFromColumn(earnCode, "EMPLOYEETYPE");

                foreach (string s in employeeType)
                {
                    cboEmployeeType.Items.Add(s.ToString().Trim());
                }

                txtPeriod.Text = Base.Period;
                cboPayRollInd.Enabled = false;
                cboEarningsCode.Enabled = false;
                cboMiningType.Enabled = false;
                cboBonusType.Enabled = false;
                cboSection.Enabled = false;

            }

        }

        private void cboPayMeth_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            cboPayrollGroup.Text = "1";  // This will only fill the box with text as ths box will be hidden and not be changed by the user any more

            if (cboEmployeeType.Text == "-- Choose one --" || cboEarningsCode.Text == "-- Choose one --"
                || cboPayRollInd.Text == "-- Choose one --" || cboTransactionTypes.Text == "-- Choose one --"
                || cboPayrollGroup.Text == "" || cboPaymentAdviceType.Text == "-- Choose one --"
                || cboPaymentUpdateType.Text == "-- Choose one --" || cboPaymentInd.Text == "-- Choose one --")
            {
                MessageBox.Show("Please ensure all boxes are filled!!", "Fill all boxes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (cboBonusName.Text == "MANUALPAY")
                {
                    btnManual_Click("me", e);
                }
                else
                {
                    txtPayPal.Visible = false;
                    this.Cursor = Cursors.WaitCursor;

                    extractReferenceNo();


                    DataTable temp = new DataTable();

                    string adviceType = cboPaymentAdviceType.Text.Substring(0, 1);

                    string strSQLCheckEntry = "Select distinct sendind from Payroll " +
                                   " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                                   "' and section = '" + cboSection.Text + "' and EARNINGSNAME = '" + cboEarningsColumnName.Text +
                                   "' and period = '" + BusinessLanguage.Period.Trim() +
                                   "' and EarningsCode  = '" + cboEarningsCode.Text.Substring(0, 2).Trim() + "'";

                    temp = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLCheckEntry);

                    if (temp.Rows.Count > 0)
                    {
                        //A previous paysend was done.
                        //Check how many rows exist.  If 2, then a "Y" and a "N" exist.  Extract the "N"'s
                        if (temp.Rows.Count > 1)
                        {
                            extractToBePaysendRecords();
                        }
                        else
                        {
                            //Only a "Y" or a "N" exists
                            string sendind = temp.Rows[0][0].ToString().Trim();
                            if (sendind == "Y")
                            {
                                //A previous paysend was done and records were send to PalPay.
                                //Create a new output table combining current records and previously payrollsend records
                                createNewPaysendTable();

                            }
                            else
                            {
                                //A previous paysend was done, but not send to PalPay yet.
                                extractToBePaysendRecords();

                            }
                        }

                    }
                    else
                    {
                        //No paysend has been done yet.
                        extractFromPayrollFile();


                    }

                    this.Cursor = Cursors.Arrow;
                }
            }
        }

        private void extractFromPayrollFile()
        {
            //Extract from payrollfile
            //if no records are found on the payroll file, then this is a new payroll send.

            string strSQLnewPayroll = "select * from PAYROLL where BUSSUNIT = '" + BusinessLanguage.BussUnit + "' and MININGTYPE = '" +
                cboMiningType.Text.Trim() + "' and BONUSTYPE = '" + cboBonusType.Text.Trim() + "' and SECTION = '" +
                cboSection.Text.Trim() + "' and REFERENCENO = '" + TB.RefNo + "' and SENDIND = 'N'";

            outputTable = TB.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLnewPayroll);

            string DBTable = cboTableName.Text;

            string strSQL;
            if (outputTable.Rows.Count == 0)
            {

                switch (DBTable)
                {

                    case "MINERS":
                        strSQL = "Select '" + BusinessLanguage.MiningType + "' as MININGTYPE,'" +
                         BusinessLanguage.BonusType + "' as BONUSTYPE," + "SECTION,EMPLOYEE_NO,MAX(GANG) AS GANG," +
                         "sum(convert(float," + cboEarningsColumnName.Text.Trim() + ")) as EARNINGSVALUE from " + DBTable + BusinessLanguage.Period +
                         " where MININGTYPE = '" + cboMiningType.Text.Trim() + "' and BONUSTYPE = '" + cboBonusType.Text.Trim() +
                         "' and SECTION = '" + cboSection.Text.Trim() +
                         "' and period = '" + BusinessLanguage.Period.Trim() +
                         "' and convert(float," + cboEarningsColumnName.Text.Trim() +
                         ") > 0 group by MININGTYPE,BONUSTYPE,SECTION,EMPLOYEE_NO,GANG";
                        break;

                    case "DRILLERSEARN":
                        strSQL = "Select '" + BusinessLanguage.MiningType + "' as MININGTYPE,'" +
                        BusinessLanguage.BonusType + "' as BONUSTYPE," + "SECTION,EMPLOYEE_NO,max(GANG) AS GANG, " +
                        "sum(convert(float," + cboEarningsColumnName.Text.Trim() + ")) as EARNINGSVALUE from " + DBTable + BusinessLanguage.Period +
                        " where MININGTYPE = '" + cboMiningType.Text.Trim() + "' and BONUSTYPE = '" + cboBonusType.Text.Trim() +
                        "' and SECTION = '" + cboSection.Text.Trim() +  
                        "' and period = '" + BusinessLanguage.Period.Trim() +
                        "' and convert(float," + cboEarningsColumnName.Text.Trim() +
                        ") <> 0 group by MININGTYPE,BONUSTYPE,SECTION,EMPLOYEE_NO";
                        break;

                    default:

                        strSQL = "Select '" + BusinessLanguage.MiningType + "' as MININGTYPE,'" +
                         BusinessLanguage.BonusType + "' as BONUSTYPE," + "SECTION,EMPLOYEE_NO,max(GANG) AS GANG, " +
                         "sum(convert(float," + cboEarningsColumnName.Text.Trim() + ")) as EARNINGSVALUE from " + DBTable + BusinessLanguage.Period +
                         " where MININGTYPE = '" + cboMiningType.Text.Trim() + "' and BONUSTYPE = '" + cboBonusType.Text.Trim() +
                         "' and SECTION = '" + cboSection.Text.Trim() +
                         "' and period = '" + BusinessLanguage.Period.Trim() +
                         "' and convert(float," + cboEarningsColumnName.Text.Trim() +
                         ") <> 0 group by MININGTYPE,BONUSTYPE,SECTION,EMPLOYEE_NO";
                        break;



                }

                //jvdw
                outputTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
                btnSendToPayrollDB.Enabled = true;
                btnDelete.Enabled = false;

                //jvdw
                dataGridView1.Columns.Clear();
                DataGridViewCheckBoxColumn c1 = new DataGridViewCheckBoxColumn();
                dataGridView1.DataSource = outputTable;
                dataGridView1.Columns.Insert(0, c1);
                chkB = new CheckBox();
                chkB.Size = new Size(15, 15);
                Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, false);

                rect.X = rect.Location.X + (rect.Width) / 2;
                rect.Y = rect.Location.Y + (rect.Height) / 4;

                dataGridView1.Controls.Add(chkB);
                chkB.CheckedChanged += new EventHandler(ckBox_CheckedChanged);

            }
            else
            {
                //btnSendToPayrollDB.Enabled = false;
                btnDelete.Enabled = true;
                btnSend.Enabled = true;
            }


        }

        private void cboEarningsCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboEarningsCode.Text.Substring(0, 2) == "--")
            {
            }
            else
            {
                if (cboEarningsCode.Text.Substring(0, 2) == "--")
                {
                }
                else
                {

                    if (cboEarningsCode.Text.Substring(0, 2) != "--")
                    {
                        string employeeType = cboEmployeeType.Text.Substring(0, 2);//Gets the entry number from the user
                        string earningsCode = cboEarningsCode.Text.ToString();
                        string getValue = "";

                        int cboPaymentIndValue = cboPayRollInd.Items.Count;//Count Values in the cboBox
                        if (cboPaymentIndValue != 0) //Checks and sets the cboBox to zero entries
                        {
                            cboPayRollInd.Items.Clear();
                        }
                        List<string> paymentInd = new List<string>();

                        getValue = "SELECT PAYROLLIND FROM EARNINGSCODES WHERE EMPLOYEETYPE = '" + cboEmployeeType.Text.ToString() + "' AND EARNINGSCODE = '" + cboEarningsCode.Text.ToString() + "'";
                        earnCode = TB.createDataTableWithAdapter(Base.DBConnectionString, getValue);
                        paymentInd = TB.loadDistinctValuesFromColumn(earnCode, "PAYROLLIND");

                        foreach (string s in paymentInd)
                        {
                            cboPayRollInd.Items.Add(s.ToString().Trim());
                        }


                        //Checks if user can choose if just one value set it as the text if not
                        //gives the user chance to choose
                        cboPaymentIndValue = cboPayRollInd.Items.Count;
                        if (cboPaymentIndValue > 1)
                        {
                            cboPayRollInd.Enabled = true;
                            cboPayRollInd.Text = "-- Choose one --";
                        }
                        else
                        {
                            cboPayRollInd.Enabled = false;
                            cboPayRollInd.Text = cboPayRollInd.Items[0].ToString();
                        }

                        lblCurrentSelectedBonus.Text = cboEarningsCode.Text;

                        //this.Cursor = Cursors.WaitCursor;
                        //extractReferenceNo();
                    }
                }
            }
        }

        private void extractReferenceNo()
        {
            if (cboMiningType.Text.Length != 0 && cboBonusType.Text.Length != 0 && cboSection.Text.Length != 0 && cboPayrollGroup.Text.Length != 0 && cboPaymentAdviceType.Text.Length != 0 && cboEarningsCode.Text.Length != 0)
            {
                //Extract all possible reference numbers that exist for this criteria.
                string strSQL = "select distinct referenceno from Payroll where bussunit = '" + BusinessLanguage.BussUnit +
                                     "' and miningtype = '" + cboMiningType.Text.Trim() +
                                     "' and bonustype = '" + cboBonusType.Text.Trim() +
                                     "' and section = '" + cboSection.Text.Trim() + "' and earningsname = ' " + cboBonusName.Text.Trim() + "' and sendind = 'N'";


                DataTable tempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                if (tempDataTable.Rows.Count > 0)
                {
                    cboReferenceNo.Items.Remove("new");

                    foreach (DataRow r in tempDataTable.Rows)
                    {
                        cboReferenceNo.Items.Add(r["ReferenceNo"].ToString().Trim());

                    }
                    cboReferenceNo.Text = tempDataTable.Rows[tempDataTable.Rows.Count - 1]["ReferenceNo"].ToString().Trim();

                }
                else
                {
                    strSQL = "select max(ReferenceNo) as LastRef from Payroll where bussunit = '" + BusinessLanguage.BussUnit +
                                         "' and miningtype = '" + cboMiningType.Text.Trim() +
                                         "' and bonustype = '" + cboBonusType.Text.Trim() +
                                         "' and section = '" + cboSection.Text.Trim() + "'";

                    tempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQL);

                    if (tempDataTable.Rows.Count > 0)
                    {
                        if ((string.IsNullOrEmpty)(tempDataTable.Rows[0][0].ToString()))
                        {
                            TB.RefNo = "new";
                            cboReferenceNo.Text = TB.RefNo.Trim();
                        }
                        else
                        {
                            TB.RefNo = "00" + Convert.ToString(Convert.ToInt16(tempDataTable.Rows[0][0].ToString()) + 1);
                            TB.RefNo = TB.RefNo.Trim().Substring(TB.RefNo.Trim().Length - 3);
                            cboReferenceNo.Text = TB.RefNo.Trim();
                        }
                    }
                    else
                    {
                        TB.RefNo = "001";
                        cboReferenceNo.Text = TB.RefNo.Trim();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please supply all input by selecting from the above comboboxes.", "Information", MessageBoxButtons.OK);
            }
            this.Cursor = Cursors.Arrow;
        }

        private void btnSendToPayrollDB_Click(object sender, EventArgs e)
        {
            if (cboSection.Text == "MANUAL")
            {

                outputTable = newPayroll;
                TB.saveCalculations2(newPayroll, Base.DBConnectionString, " where period = '999999'", "Payroll");

                string strSQLpreviousPaysend = "Select * from Payroll " +
                                   " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                                   "' and section = 'MANUAL' and EARNINGSNAME  = '" + cboBonusName.Text.Trim() +
                                   "' and period = '" + BusinessLanguage.Period.Trim() +
                                   "' and REFERENCENO = '" + TB.RefNo + "' order by miningtype,bonustype,section,Employee_No";

                outputTable = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLpreviousPaysend);
                dataGridView1.DataSource = outputTable;

                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {

                    for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                    {

                        dataGridView1[j, i].Style.BackColor = Color.Lavender;
                    }
                }

                this.dataGridView1.Refresh();
            }

            else
            {
                #region smp100
                string earningcode = "";
                this.Cursor = Cursors.WaitCursor;
                string dateTimeNow = DateTime.Now.ToString();
                StringBuilder strSQL = new StringBuilder("BEGIN transaction;");
                string strPayrollReturn = string.Empty;

                if (cboReferenceNo.Text.Trim() == "new")
                {
                    TB.RefNo = "001";

                }

                int enterValue = 0;


                foreach (DataRow r in outputTable.Rows)
                {

                    string test2 = Convert.ToString(decimal.Round(Convert.ToDecimal(r["EARNINGSVALUE"].ToString().Trim()), 2)).Replace(".", "");
                    if (test2 != "0")
                    {

                        enterValue = 1;
                        if (decimal.Round(Convert.ToDecimal(r["EARNINGSVALUE"].ToString().Trim()), 2) < 0)
                        {
                            string strSign = string.Empty;
                            if (Convert.ToDecimal(r["EARNINGSVALUE"].ToString().Trim()) > 0)
                            {
                                strSign = "+";
                            }
                            else
                            {
                                strSign = "-";
                            }
                            string valueEarning = r["earningsvalue"].ToString().Trim();
                            int comaPlace = valueEarning.LastIndexOf(".");
                            string sub = valueEarning.Substring(comaPlace + 1).ToString();
                            int placesAfterComma = valueEarning.Substring(comaPlace + 1).ToString().Length;


                            if (placesAfterComma != 2 || comaPlace < 0)
                            {
                                if (placesAfterComma == 1)
                                {
                                    valueEarning = valueEarning + "0";
                                }
                                else
                                {
                                    valueEarning = valueEarning + ".00";
                                }
                            }

                            DateTime today = DateTime.Today;

                            strSQL.Append("insert into payroll values('" + BusinessLanguage.BussUnit +
                                                      "','" + cboMiningType.Text.Trim() + "','" + cboBonusType.Text.Trim() +
                                                      "','" + cboSection.Text.Trim() + "','" + r["GANG"].ToString().Trim() + "','" + BusinessLanguage.Period +
                                                      "','" + TB.RefNo + "','" + cboBonusName.Text.Trim() +
                                                      "','" + cboTableName.Text.Trim() + "','" +
                                                      r["EMPLOYEE_NO"].ToString().Trim() + "','" +
                                                      cboEarningsCode.Text.Substring(0, 2).Trim() + "','" +
                                                      cboPaymentInd.Text.Trim().Substring(0, cboPaymentInd.Text.Trim().IndexOf("-")) +
                                                      "','" + today.ToShortDateString() +
                                                      "','" + Convert.ToString(decimal.Round(Convert.ToDecimal(valueEarning.ToString()), 2)) + "','" + strSign + "','" +
                                                      dateTimeNow.ToString() + "','N','Not Send');  ");
                        }
                        else
                        {
                            DateTime today = DateTime.Today;
                            string strSign = string.Empty;
                            if (Convert.ToDecimal(r["earningsvalue"].ToString().Trim()) > 0)
                            {
                                strSign = "+";
                            }
                            else
                            {
                                strSign = "-";
                            }

                            string test = TB.RefNo;

                            string valueEarning = r["earningsvalue"].ToString().Trim();
                            int comaPlace = valueEarning.LastIndexOf(".");
                            string sub = valueEarning.Substring(comaPlace + 1).ToString();
                            int placesAfterComma = valueEarning.Substring(comaPlace + 1).ToString().Length;


                            if (placesAfterComma != 2 || comaPlace < 0)
                            {
                                if (placesAfterComma == 1)
                                {
                                    valueEarning = valueEarning + "0";
                                }
                                else
                                {
                                    valueEarning = valueEarning + ".00";
                                }
                            }

                            test = Convert.ToString(decimal.Round(Convert.ToDecimal(valueEarning), 2)).Replace(".", "");

                            strSQL.Append("insert into payroll values('" + BusinessLanguage.BussUnit +
                                "','" + cboMiningType.Text.Trim() + "','" + cboBonusType.Text.Trim() +
                                "','" + cboSection.Text.Trim() + "','" + r["GANG"].ToString().Trim() + "','" + BusinessLanguage.Period +
                                "','" + TB.RefNo + "','" + cboBonusName.Text.Trim() +
                                "','" + cboTableName.Text.Trim() + "','" +
                                r["EMPLOYEE_NO"].ToString().Trim() + "','" +
                                cboEarningsCode.Text.Substring(0, 2).Trim() + "','" +
                                cboPaymentInd.Text.Trim().Substring(0, cboPaymentInd.Text.Trim().IndexOf("-")) +
                                "','" + today.ToShortDateString() +
                                "','" + valueEarning + "','" + strSign + "','" +
                                dateTimeNow.ToString() + "','N','Not Send');  ");

                        }
                    }
                }

                strSQL.Append(" COMMIT transaction;");

                if (enterValue == 1)
                {
                    strPayrollReturn = TB.InsertPayrollData(Base.DBConnectionString.ToString(), strSQL.ToString());



                    if (strPayrollReturn.Contains("Violation of PRIMARY KEY"))
                    {
                        MessageBox.Show("These employees have been paysend previously.", "Information", MessageBoxButtons.OK);
                        //Extract records that are on the payroll table, but has not yet been paysend.

                        string strSQLtest1 = "Select * from Payroll " +
                                   " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                                   "' and section = '" + cboSection.Text + "' and EarningsCode  = '" + earningcode.Trim() +
                                   "' and sendind = 'N'" +
                                   "' and period = '" + BusinessLanguage.Period.Trim() +
                                   " order by miningtype,bonustype,section,Employee_No";

                        outputTable = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLtest1);
                        //outputTable = Base.extractPreviousPaysend(Base.DBConnectionString.ToString(), cboMiningType.Text.Trim(), cboBonusType.Text.Trim(),
                        //cboSection.Text.Trim(), cboPaymentAdviceType.Text.Trim().Substring(0, 1),
                        //cboEarningsCode.Text.Trim(), "N");
                        if (outputTable.Rows.Count > 0)
                        {
                            btnDelete.Enabled = true;
                            btnSend.Enabled = true;
                            btnSendToPayrollDB.Enabled = false;
                            dataGridView1.DataSource = outputTable;
                            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                            {

                                for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                                {

                                    dataGridView1[j, i].Style.BackColor = Color.Lavender;
                                }
                            }

                            this.dataGridView1.Refresh();
                        }
                        else
                        {
                            string strSQLtest = "select miningtype,bonustype,section,Employee_No,EarningsCode,sum(convert(float,BonusValue)) as BonusValue, " +
                  " sum(convert(float,PayrollValue)) as PayrollValue,sum(convert(float,BonusValue))+ sum(convert(float,PayrollValue)) as EarningsValue " +
                  " from(" +
                  " Select miningtype, " +
                  " bonustype,section,Employee_No,EarningsCode, " +
                  " sum(convert(float,EmployeeRiggingBonus)) as BonusValue, 0 as PayrollValue from BONUSSHIFTS  " +
                  " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                  "' and period = '" + BusinessLanguage.Period.Trim() +
                  "' and section = '" + cboSection.Text + "' " +
                  " and EarningsCode = '" + "KA" + "' and convert(float,EmployeeRiggingBonus) > 0  " +
                  " group by  miningtype,bonustype,section,Employee_No,EarningsCode" +
                  " union" +
                  " select miningtype,bonustype,section,Employee_No,EarningsCode,0 as BonusValue,sum(convert(float,BonusValue)*-1) as PayrollValue " +
                  " from payroll where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                  "' and period = '" + BusinessLanguage.Period.Trim() +
                  "' and section = '" + cboSection.Text + "' " +
                  " and EarningsCode = '" + "KA" + "' and sendind = 'Y'" +
                  " group by  miningtype,bonustype,section,Employee_No,EarningsCode) as t1 " +
                  " group by  miningtype,bonustype,section,Employee_No,EarningsCode";
                            //outputTable = Base.createNewPaysend(BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                            //cboSection.Text.Trim(), cboEarningsCode.Text.Trim(), Base.DBConnectionString.ToString());
                            btnDelete.Enabled = false;
                            btnSend.Enabled = false;
                            btnSendToPayrollDB.Enabled = false;
                            dataGridView1.DataSource = outputTable;
                            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                            {

                                for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                                {

                                    dataGridView1[j, i].Style.BackColor = Color.Yellow;
                                }
                            }

                            this.dataGridView1.Refresh();
                        }
                    }
                    else
                    {


                        string strSQLpreviousPaysend = "Select * from Payroll " +
                                   " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                                   "' and section = '" + cboSection.Text + "' and EARNINGSNAME  = '" + cboBonusName.Text.Trim() +
                                   "' and period = '" + BusinessLanguage.Period.Trim() +
                                   "'and REFERENCENO = '" + TB.RefNo + "' order by miningtype,bonustype,section,Employee_No";

                        outputTable = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLpreviousPaysend);

                        //Base.extractPreviousPaysend(Base.DBConnectionString.ToString(), cboMiningType.Text.Trim(), cboBonusType.Text.Trim(),
                        // cboSection.Text.Trim(), cboPaymentAdviceType.Text.Trim().Substring(0, 1),
                        //cboEarningsCode.Text.Trim(), "N");
                        dataGridView1.DataSource = outputTable;

                        for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                        {

                            for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                            {

                                dataGridView1[j, i].Style.BackColor = Color.Lavender;
                            }
                        }
                        this.dataGridView1.Refresh();
                    }

                    dataGridView1.DataSource = outputTable;

                    MessageBox.Show("Records are ready to be send to PalPay", "Information", MessageBoxButtons.OK);
                    btnSendToPayrollDB.Enabled = false;
                    btnSend.Enabled = true;
                }
                else
                {
                    MessageBox.Show("No new data to add!!", "Information", MessageBoxButtons.OK);
                    btnSendToPayrollDB.Enabled = false;
                    //btnNewPayRoll.Enabled = true;

                }
                this.Cursor = Cursors.Arrow;

                #endregion
            }
        }

        private void cboReferenceNo_SelectedIndexChanged(object sender, EventArgs e)
        {


            this.Cursor = Cursors.Arrow;
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            int counterpositve = 0;
            int counternegative = 0;
            DataTable positiveTable = new DataTable();
            DataTable negativeTable = new DataTable();
            positiveTable = outputTable.Clone();
            negativeTable = outputTable.Clone();

            foreach (DataRow r in outputTable.Rows)
            {
                if (r["earningssign"].ToString().Trim() == "+")
                {

                    positiveTable.Rows.Add(r.ItemArray);
                    counterpositve = counterpositve + 1;
                }
                if (r["earningssign"].ToString().Trim() == "-")
                {
                    negativeTable.Rows.Add(r.ItemArray);
                    counternegative = counternegative + 1;
                }

            }
            if (counterpositve > 0)
            {
                string answer = MessageBox.Show("Do you want to paysend earnings?", "Earnings", MessageBoxButtons.YesNo, MessageBoxIcon.Question).ToString();
                if (answer == "Yes")
                {
                    textFilePayroll(positiveTable);
                }
            }
            if (counternegative > 0)
            {
                string answer = MessageBox.Show("Do you want to paysend deduction?", "Deductions", MessageBoxButtons.YesNo, MessageBoxIcon.Question).ToString();
                if (answer == "Yes")
                {
                    textFilePayroll(negativeTable);
                }
            }
            if (cboReferenceNo.Text == "new")
            {
                outputTable = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "PAYROLL", " Where referenceno = '001'" +
                                                                 " and section = '" + cboSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period.Trim() + "'");
            }
            else
            {
                outputTable = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "PAYROLL", " Where referenceno = '" + cboReferenceNo.Text.Trim() +
                                                                 "' and section = '" + cboSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period.Trim() + "'");
            }

            //btnNewPayRoll.Enabled = true;
            btnSend.Enabled = false;
            cboEmployeeType.Enabled = false;

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete the payroll records?", "Confirmation", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:

                    string strSQL = "Delete from Payroll where bussunit = '" + BusinessLanguage.BussUnit +
                                   "' and miningtype = '" + cboMiningType.Text.Trim() +
                                   "' and bonustype = '" + cboBonusType.Text.Trim() +
                                   "' and section = '" + cboSection.Text.Trim() +
                                   "' and Earningscode = '" + cboEarningsCode.Text.Substring(0, 2).Trim() +
                                   "' and period = '" + BusinessLanguage.Period.Trim() +
                                   "' and SendInd = 'N'";


                    TB.InsertData(Base.DBConnectionString, strSQL.ToString());

                    MessageBox.Show("Records removed from the payroll database.", "Info", MessageBoxButtons.OK);

                    btnGenerate_Click(sender, e);
                    //extractFromPayrollFile();

                    break;

                case DialogResult.No:
                    break;

            }

        }

        private void cboPayrollGroup_SelectedIndexChanged(object sender, EventArgs e)
        {

            DataTable transactiontypes = Base.getTransActionType(Base.DBConnectionString);

            foreach (DataRow x in transactiontypes.Rows)
            {
                cboTransactionTypes.Items.Add(x["TransactionType"].ToString().Trim() + "-" + x["TransactionDesc"].ToString().Trim());
            }
            if (transactiontypes.Rows.Count == 1)
            {
                cboTransactionTypes.Text = transactiontypes.Rows[0]["TransactionType"].ToString().Trim()
                                         + "-" + transactiontypes.Rows[0]["TransactionDesc"].ToString().Trim();

            }

            DataTable advicetypes = Base.getAdviceTypes(Base.DBConnectionString);
            foreach (DataRow x in advicetypes.Rows)
            {
                cboPaymentAdviceType.Items.Add(x["PaymentAdviceType"].ToString().Trim() + "-" + x["PaymentAdviceDesc"].ToString().Trim());
            }

            if (advicetypes.Rows.Count == 1)
            {
                cboPaymentAdviceType.Text = advicetypes.Rows[0]["PaymentAdviceType"].ToString().Trim()
                                         + "-" + advicetypes.Rows[0]["PaymentAdviceDesc"].ToString().Trim();

            }

            DataTable paymentInd = Base.getPaymentInd(Base.DBConnectionString);
            foreach (DataRow x in paymentInd.Rows)
            {
                cboPayRollInd.Items.Add(x["PaymentInd"].ToString().Trim() + "-" + x["PaymentIndDesc"].ToString().Trim());
            }

            if (paymentInd.Rows.Count == 1)
            {
                cboPayRollInd.Text = paymentInd.Rows[0]["PaymentInd"].ToString().Trim()
                                   + "-" + paymentInd.Rows[0]["PaymentIndDesc"].ToString().Trim();

            }

            DataTable paymentUpdateType = Base.getPaymentUpdateTypes(Base.DBConnectionString);
            foreach (DataRow x in paymentUpdateType.Rows)
            {
                cboPaymentUpdateType.Items.Add(x["PaymentUpdateType"].ToString().Trim() + "-" + x["PaymentUpdateDesc"].ToString().Trim());
            }

            if (paymentUpdateType.Rows.Count == 1)
            {
                cboPaymentUpdateType.Text = paymentUpdateType.Rows[0]["PaymentUpdateType"].ToString().Trim()
                                          + "-" + paymentUpdateType.Rows[0]["PaymentUpdateDesc"].ToString().Trim();

            }



            //Extract the possibile earnings codes according to the payrollgroup and payrollsubtypes selected by the user.
            //test = cboPayrollSubTypes.Text.Trim();
            string test2 = cboPayrollGroup.Text.Trim().Substring(0, cboPayrollGroup.Text.IndexOf("-"));

            string strSQL = "SELECT distinct EARNINGSCODE FROM dbo.EarningsCode " +
                            "where PayrollGroup = '" + test2 + "'";

            SqlConnection tst = BaseConn;
            SqlCommand tstCommand = new SqlCommand(strSQL, tst);

            SqlDataAdapter tstAdapter = new SqlDataAdapter(tstCommand);
            DataTable testTBB = new DataTable();

            tstAdapter.Fill(testTBB);

            if (testTBB.Rows.Count > 1)
            {
                foreach (DataRow r in testTBB.Rows)
                {
                    cboEarningsCode.Items.Add(r["EARNINGSCODE"].ToString().Trim() + "-" + r["EARNINGSCODE_DESC"].ToString().Trim());
                    cboEarningsCode.Enabled = true;
                }
            }
            else
            {
                if (testTBB.Rows.Count == 1)
                {
                    cboEarningsCode.Text = testTBB.Rows[0]["EARNINGSCODE"].ToString().Trim();
                    cboEarningsCode.Enabled = false;

                }
            }
            //==========================================================================================================

            //Update the reference combobox

            extractReferenceNo();

            //==========================================================================================================

            //Color and enable all the boxes necesarry

            cboPaymentAdviceType.Enabled = true;
            cboReferenceNo.Enabled = false;
            cboEarningsColumnName.Enabled = true;

            cboReferenceNo.BackColor = Color.Lavender;
            cboEarningsCode.BackColor = Color.Lavender;
            cboEarningsColumnName.BackColor = Color.Lavender;
            cboPaymentAdviceType.BackColor = Color.Lavender;
            cboPaymentUpdateType.BackColor = Color.Lavender;


        }

        private void cboEmployeeType_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboEarningsCode.Items.Clear();
            cboEarningsCode.Text = "";
            cboPayRollInd.Items.Clear();
            cboPayRollInd.Text = "";
            cboBonusName.Items.Clear();
            cboBonusName.Text = "";
            cboTableName.Items.Clear();
            cboTableName.Text = "";
            cboEarningsColumnName.Text = "";


            if (cboEmployeeType.Text.Substring(0, 2) != "--")
            {
                string employeeType = cboEmployeeType.Text.Substring(0, 2);//Gets the entry number from the user


                int cboCodesValue = cboEarningsCode.Items.Count;//Count Values in the cboBox
                if (cboCodesValue != 0) //Checks and sets the cboBox to zero entries
                {
                    cboEarningsCode.Items.Clear();
                }
                List<string> earningsCode = new List<string>();
                //switch (employeeType)//Reads the entry from the user
                //{
                ////Officials
                //case "1 ":
                string getValueEmployType = "SELECT EARNINGSCODE FROM EARNINGSCODES WHERE EMPLOYEETYPE = '" + cboEmployeeType.Text.ToString() +
                                            "'and bonustype = '" + cboBonusType.Text.ToString() + "'";
                earnCode = TB.createDataTableWithAdapter(Base.DBConnectionString, getValueEmployType);
                earningsCode = TB.loadDistinctValuesFromColumn(earnCode, "EARNINGSCODE");

                foreach (string s in earningsCode)
                {
                    cboEarningsCode.Items.Add(s.ToString().Trim());
                }


                //Checks if user can choose if just one value set it as the text if not
                //gives the user chance to choose
                cboCodesValue = cboEarningsCode.Items.Count;
                if (cboCodesValue > 1)
                {
                    cboEarningsCode.Enabled = true;
                    cboEarningsCode.Text = "-- Choose one --";
                }
                else
                {
                    cboEarningsCode.Enabled = false;
                    cboEarningsCode.Text = cboEarningsCode.Items[0].ToString();
                }

            }
            else
            {
                MessageBox.Show("Must choose one!!");
            }
        }

        //This check and uncheck all checkboxes on the screen
        void ckBox_CheckedChanged(object sender, EventArgs e)
        {

            for (int j = 0; j < this.dataGridView1.RowCount; j++)
            {

                this.dataGridView1[0, j].Value = this.chkB.Checked;

            }

            this.dataGridView1.EndEdit();

        }

        private void cboPaymentInd_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboTableName.Items.Clear();
            cboBonusName.Items.Clear();

            List<string> tableName = new List<string>();
            List<string> bonusName = new List<string>();
            string getValuePayInd = string.Empty;
            //AMP
            if (cboSection.Text.Trim() == "MANUAL")
            {
                getValuePayInd = "SELECT TABLENAME, BONUSNAME FROM EARNINGSCODES WHERE EMPLOYEETYPE = '" + cboEmployeeType.Text.ToString() +
                                        "' AND EARNINGSCODE = '" + cboEarningsCode.Text.ToString() +
                                        "' AND PAYROLLIND= '" + cboPayRollInd.Text.ToString() +
                                        "' and bonustype = '" + cboBonusType.Text.ToString() + "' AND TABLENAME = 'MANUALPAY' ";
            }
            else
            {
                getValuePayInd = "SELECT TABLENAME, BONUSNAME FROM EARNINGSCODES WHERE EMPLOYEETYPE = '" + cboEmployeeType.Text.ToString() +
                                        "' AND EARNINGSCODE = '" + cboEarningsCode.Text.ToString() +
                                        "' AND PAYROLLIND= '" + cboPayRollInd.Text.ToString() +
                                        "' and bonustype = '" + cboBonusType.Text.ToString() + "' AND TABLENAME != 'MANUALPAY' ";
            }

            earnCode = TB.createDataTableWithAdapter(Base.DBConnectionString, getValuePayInd);
            tableName = TB.loadDistinctValuesFromColumn(earnCode, "TABLENAME");
            bonusName = TB.loadDistinctValuesFromColumn(earnCode, "BONUSNAME");

            foreach (string s in tableName)
            {
                cboTableName.Items.Add(s.ToString().Trim());
            }

            foreach (string s in bonusName)
            {
                cboBonusName.Items.Add(s.ToString().Trim());
            }

            int cboTableNameCount = cboTableName.Items.Count;
            int cboBonusNameCount = cboBonusName.Items.Count;

            if (cboTableNameCount > 1)
            {
                cboTableName.Enabled = true;
                cboTableName.Text = "-- Choose one --";
            }
            else
            {
                cboTableName.Enabled = false;
                cboTableName.Text = cboTableName.Items[0].ToString();
            }

            if (cboBonusNameCount > 1)
            {
                cboBonusName.Enabled = true;
                cboBonusName.Text = "-- Choose one --";
            }
            else
            {
                cboBonusName.Enabled = false;
                cboBonusName.Text = cboBonusName.Items[0].ToString();
            }



        }

        private void cboBonusName_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboEarningsColumnName.Text = cboBonusName.Text;
        }

        #region vanaf stmp1000
        //VANAF STMP1000

        private void extractToBePaysendRecords()
        {
            string strSQLPrevious = "Select * from Payroll " +
                                 " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                                 "' and section = '" + cboSection.Text + "' and EarningsCode  = '" + cboEarningsCode.Text.Substring(0, 2) +
                                 "' and sendind = '" + "N" +
                                 "' and period = '" + BusinessLanguage.Period.Trim() +
                                 "' order by miningtype,bonustype,section,Employee_No";

            outputTable = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLPrevious);

            // Base.extractPreviousPaysend(Base.DBConnectionString.ToString(), BusinessLanguage.MiningType, BusinessLanguage.BonusType,
            //    cboSection.Text.Trim(), cboPaymentAdviceType.Text.Trim().Substring(0, 1),
            //   cboEarningsCode.Text.Trim(), "N");
            //Display on table and make it blue
            dataGridView1.DataSource = outputTable;
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {

                for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                {

                    dataGridView1[j, i].Style.BackColor = Color.Aqua;
                }
            }

            this.dataGridView1.Refresh();
            btnDelete.Enabled = true;
            btnSendToPayrollDB.Enabled = false;
            btnSend.Enabled = true;

        }
        private void createNewPaysendTable()
        {
            btnDelete.Enabled = false;
            btnSend.Enabled = false;
            btnSendToPayrollDB.Enabled = true;

            string sqlStringNewPaysend = "select miningtype,bonustype,section,max(GANG) AS GANG,Employee_No," +
                                          "round(sum(convert(float,PayrollValue)),2) as PrevPaysend,  " +
                                          " round(sum(convert(float,BonusValue)),2) as NewRun,round(sum(convert(float,BonusValue))-" +
                                          "sum(convert(float,PayrollValue)),2) as EARNINGSVALUE " +
                                          "from(" +
                                            "Select miningtype,  bonustype,section,max(GANG)AS GANG,Employee_No, " +
                                            " sum(convert(float," + cboBonusName.Text.Trim() + ")) as BonusValue, 0 as PayrollValue" +
                                            " from " + cboTableName.Text.Trim() + BusinessLanguage.Period.Trim() + " " +
                                            " where miningtype = '" + cboMiningType.Text.Trim() + "' and bonustype = '" + cboBonusType.Text.Trim() +
                                            "' and section = '" + cboSection.Text.Trim() +
                                            "' and period = '" + BusinessLanguage.Period.Trim() +
                                            "' and convert(float," + cboBonusName.Text.Trim() +
                                            ") <> 0   group by  miningtype,bonustype,section," + "Employee_No " +

                                         "union select miningtype,bonustype,section,max(GANG)AS GANG,Employee_No,0 as BonusValue," +
                                            " sum(convert(float,EARNINGSVALUE)) as PayrollValue  " +
                                            " from payroll where miningtype = '" + cboMiningType.Text.Trim() + "'" +
                                            " and bonustype = '" + cboBonusType.Text.Trim() + "' and section = '" + cboSection.Text.Trim() +
                                            "' and period = '" + BusinessLanguage.Period.Trim() + "'" +
                                            " and EarningsCode = '" + cboEarningsCode.Text.Substring(0, 2).Trim() +
                                            "' and sendind = 'Y' AND EARNINGSNAME = '" + cboBonusName.Text.Trim() + "'" +
                                            "group by  miningtype,bonustype,section,Employee_No,EarningsCode) as t1 " +
                                            "group by  miningtype,bonustype,section,Employee_No";

            outputTable = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), sqlStringNewPaysend);

            //Base.createNewPaysend(BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
            //cboSection.Text.Trim(), cboEarningsCode.Text.Trim(), Base.DBConnection.ToString());

            dataGridView1.DataSource = outputTable;
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {

                for (int j = 0; j <= dataGridView1.Columns.Count - 1; j++)
                {

                    dataGridView1[j, i].Style.BackColor = Color.LightYellow;
                }
            }
            this.dataGridView1.Refresh();
            btnDelete.Enabled = false;
            btnSendToPayrollDB.Enabled = true;
            btnSend.Enabled = false;
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void textFilePayroll(DataTable workingtable)
        {
            string payIndicator = "";
            if (cboPaymentInd.Text.Substring(0, 1).Trim() == "*")
            {
                payIndicator = " ";

            }
            else
            {
                payIndicator = cboPaymentInd.Text.Substring(0, 1).Trim();
            }
            //Send to text file
            #region from smtp1000 jvdw

            //Create Batchnumber 
            string batchNrString = "";
            string strSQLGetBatch = "SELECT max(convert(int,BATCH_NO)) as lastBatch FROM PAYROLLCOUNTER";
            DataTable batchNoTable = new DataTable();
            batchNoTable = Base.createDataTableWithAdapter(Base.BaseConnectionString.ToString(), strSQLGetBatch);
            batchNrString = batchNoTable.Rows[0][0].ToString().Trim();
            if (batchNoTable.Rows[0][0].ToString() == "" || batchNoTable.Rows[0][0].ToString() == "00" || batchNoTable.Rows[0][0].ToString().Trim() == "99")
            {
                if (batchNoTable.Rows[0][0].ToString().Trim() == "99")
                {
                    DataTable payCountAchive = new DataTable();
                    payCountAchive = TB.createDataTableWithAdapterSelectAll(Base.BaseConnectionString, "PAYROLLCOUNTER");
                    TB.saveCalculations2(payCountAchive, Base.BaseConnectionString, "WHERE PERIOD = '" + txtPeriod.Text.ToString().Trim() + "'", "PAYROLLCOUNTERACHIVE");
                    TB.deleteAllExcept(Base.BaseConnectionString, "PAYROLLCOUNTER");


                }
                batchNrString = "00";
            }
            int batchNrInt = Convert.ToInt16(batchNrString) + 1;

            string databaseBatch = Convert.ToString(batchNrInt);


            if (batchNrInt.ToString().Length < 2)
            {
                batchNrString = Convert.ToString(batchNrInt).Insert(0, "0");
            }
            else
            {
                batchNrString = batchNrInt.ToString().Trim();

            }

            string paymeth = "";
            decimal earningsTotal = 0;
            string EarningsSign = string.Empty;

            //string strCount = "";

            //if (cboPaymentUpdateType.Text.Trim().Substring(0, 1) == "M")
            //{
            //    paymeth = "1";
            //}
            //else
            //{

            //    paymeth = cboPaymentInd.Text.Substring(0, cboPaymentUpdateType.Text.IndexOf("-"));
            //}
            bool check = Directory.Exists("c:\\iCalc PayFiles\\" + Base.Period + "\\Stoping");
            if (check != true)
            {
                Directory.CreateDirectory(("c:\\iCalc PayFiles\\" + Base.Period + "\\Stoping"));
            }

            DateTime today = DateTime.Today;
            string OPath = "c:\\iCalc PayFiles\\" + Base.Period + "\\Stoping\\" + cboPayRollInd.Text.Substring(0, 2) + txtBussUnit.Text.Substring(0, 2).Trim() + today.ToString("MMdd") + "0" + batchNrString + "B.PAY";
            int entryCount = 0; //Sets the counter of record to 0;


            StreamWriter sw = new StreamWriter(OPath);
            if (workingtable.Rows.Count > 0)
            {

                //sw.WriteLine("                                    " + OPath.Remove(0, 3));

                foreach (DataRow r in workingtable.Rows)
                {
                    decimal earningsValue = decimal.Round(Convert.ToDecimal(r["earningsvalue"].ToString()), 2);
                    earningsTotal = earningsTotal + Convert.ToDecimal(earningsValue);
                    string strdte = r["senddate"].ToString().Substring(6, 2) + r["senddate"].ToString().Substring(4, 2) +
                                    r["senddate"].ToString().Substring(2, 2);

                    if (earningsValue < 0)
                    {
                        EarningsSign = "-";

                    }
                    if (r["EARNINGSSIGN"].ToString().Trim() == "+")
                    {
                        EarningsSign = " ";
                    }

                    string strEarnings = Convert.ToString(Convert.ToDecimal(r["earningsvalue"].ToString().Replace(".", "")));
                    if (strEarnings[0].ToString() == "-")
                    {
                        strEarnings = strEarnings.Remove(0, 1);
                    }
                    string strnulls = "00000000000";

                    strEarnings = strnulls.Substring(0, 11 - strEarnings.Length) + strEarnings.Trim();

                    sw.WriteLine(cboPayRollInd.Text.Substring(0, 2).Trim() + r["EMPLOYEE_NO"].ToString().Trim() + "    " + cboTransactionTypes.Text.Substring(0, 1).Trim() + r["EARNINGSCODE"].ToString().Trim() + " " + cboPaymentUpdateType.Text.Substring(0, 2).Trim() + payIndicator.Trim() + "      " + strEarnings + EarningsSign.ToString().Trim());
                    entryCount = entryCount + 1;
                }
                string entryCountString = Convert.ToString(entryCount);

                while (entryCountString.Length != 4)
                {
                    entryCountString = entryCountString.Insert(0, "0");
                }

                string strEarningsTotal = Convert.ToString(Convert.ToInt32(earningsTotal * 100));
                if (earningsTotal < 0)
                {
                    EarningsSign = "-";
                    strEarningsTotal = strEarningsTotal.Remove(0, 1);

                }
                while (strEarningsTotal.Length != 15)
                {
                    strEarningsTotal = strEarningsTotal.Insert(0, "0");
                }


                sw.WriteLine(cboPayRollInd.Text.Substring(0, 2).Trim() + "************ZZJP00" + batchNrString + entryCountString + strEarningsTotal.Trim() + EarningsSign);

                //if (outputTable.Rows.Count < 100)
                //{
                //    strCount = "0" + Convert.ToString(outputTable.Rows.Count).Trim();
                //}

                //sw.WriteLine("SI", Employee_No, TransactionType, earningscode, paymentadvicetype, paymentupdatetype, paymentind, 
                //senddate, earningsvalue, " ");
                sw.Close();

                //Base.UpdatePayrollTable(strDay, strMonth, BusinessLanguage.BussUnit, cboMiningType.Text.Trim(), cboBonusType.Text.Trim(),
                //cboSection.Text.Trim(), TB.TBName.Trim(), cboPaymentAdviceType.Text.Trim().Substring(0, 1),
                //cboReferenceNo.Text.Trim(), cboEarningsCode.Text.Trim(), "N", Base.DBConnectionString.ToString());

                string dateTimeNow = DateTime.Now.ToString();
                if (EarningsSign == " ")
                {
                    EarningsSign = "+";
                }
                string updateStrSQL = "update Payroll  set SendInd = 'Y', SendDate = '" + dateTimeNow + "', BATCHNO = '" + batchNrString +
                                   "' where bussunit = '" + BusinessLanguage.BussUnit +
                                   "' and miningtype = '" + cboMiningType.Text.Trim() +
                                   "' and bonustype = '" + cboBonusType.Text.Trim() +
                                   "' and section = '" + cboSection.Text.Trim() +
                                   "' and EARNINGSNAME = '" + cboBonusName.Text.ToString().Trim() +
                                   "' and EarningsCode = '" + cboEarningsCode.Text.Substring(0, 2).Trim() + "' and SendInd = '" + "N" + "' and earningssign = '" + EarningsSign + "'";


                string strSQLLogPayroll = "insert into PAYROLLCOUNTER values ('" + cboMiningType.Text + "','" + cboBonusType.Text.Trim() + "','" + txtPeriod.Text.Trim() + "','" + cboSection.Text.Trim() + "'," + "'JB','" + databaseBatch.Trim() + "','" + dateTimeNow + "','" + earningsTotal + "','" + Base.Userid.ToString().Trim() + "','" + TB.RefNo + "')";

                if (cboBonusName.Text == "MANUALPAY")
                {

                    string updateManualTableSQL = "update MANUALPAY  set SendInd = 'Y'";
                    Base.VoidQuery(Base.DBConnectionString, updateManualTableSQL);
                }


                Base.VoidQuery(Base.DBConnectionString, updateStrSQL);
                Base.VoidQuery(Base.BaseConnectionString, strSQLLogPayroll);


                MessageBox.Show("Textfile: " + OPath + " was created", "Information", MessageBoxButtons.OK);

                outputTable = Base.extractPreviousPaysend(Base.DBConnectionString.ToString(), BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                                                          cboSection.Text.Trim(), cboPaymentAdviceType.Text.Trim().Substring(0, 1),
                                                          cboEarningsCode.Text.Trim());
                dataGridView1.DataSource = outputTable;

                txtPayPal.LoadFile(OPath, RichTextBoxStreamType.PlainText);
                txtPayPal.Visible = true;
                txtPayPal.BackColor = Color.Lavender;



            }
            else
            {
                MessageBox.Show("No records available for payroll send", "Information", MessageBoxButtons.OK);
            }

            #endregion

        }

        private void btnNewPayRoll_Click(object sender, EventArgs e)
        {
            string answer = MessageBox.Show("Do you want to paysend a new SECTION??", "Question?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question).ToString();

            switch (answer)
            {
                case "Yes":
                    MessageBox.Show("Please choose a new section to Paysend!", "Question?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    scrLogon.ActiveForm.Close();
                    break;

                case "No":
                    MessageBox.Show("Please choose a new Earnings Code!", "Question", MessageBoxButtons.OK, MessageBoxIcon.Question);

                    break;

            }

            if (cboEarningsCode.Items.Count > 2)
            {
            }
            if (cboEarningsCode.Items.Count == 1)
            {
                txtPayPal.Visible = false;
            }
        }

        private void txtPeriod_TextChanged(object sender, EventArgs e)
        {
            lblPeriod2.Text = txtPeriod.Text + " - " + cboSection.Text.Trim();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            scrPayroll.ActiveForm.Close();
        }

        private void btnManual_Click(object sender, EventArgs e)//JVDW
        {
            DataTable temp = new DataTable();

            cboPayrollGroup.Text = "1";
            if (cboEmployeeType.Text == "-- Choose one --" || cboEarningsCode.Text == "-- Choose one --"
                || cboPayRollInd.Text == "-- Choose one --" || cboTransactionTypes.Text == "-- Choose one --"
                || cboPayrollGroup.Text == "" || cboPaymentAdviceType.Text == "-- Choose one --"
                || cboPaymentUpdateType.Text == "-- Choose one --" || cboPaymentInd.Text == "-- Choose one --")
            {
                MessageBox.Show("Please ensure all boxes are filled!!", "Fill all boxes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                extractReferenceNo();
                string adviceType = cboPaymentAdviceType.Text.Substring(0, 1);

                string strSQLCheckEntry = "Select distinct sendind from Payroll " +
                               " where miningtype = '" + cboMiningType.Text + "' and bonustype = '" + cboBonusType.Text +
                               "' and section = 'MANUAL' and EARNINGScode = '" + cboEarningsCode.Text.Trim().Substring(0, cboEarningsCode.Text.Trim().IndexOf("-") - 1) +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

                temp = Base.createDataTableWithAdapter(Base.DBConnectionString.ToString(), strSQLCheckEntry);

                if (temp.Rows.Count > 0)
                {
                    //A previous paysend was done.
                    //Check how many rows exist.  If 2, then a "Y" and a "N" exist.  Extract the "N"'s
                    if (temp.Rows.Count > 1)
                    {
                        extractToBePaysendRecords();
                    }
                    else
                    {
                        //Only a "Y" or a "N" exists
                        string sendind = temp.Rows[0][0].ToString().Trim();
                        if (sendind == "Y")
                        {
                            //A previous paysend was done and records were send to PalPay.
                            //Create a new output table combining current records and previously payrollsend records
                            //createNewPaysendTable();

                            manualPay = new DataTable();

                            manualPay = TB.createDataTableWithAdapter(Base.DBConnectionString, "SELECT * FROM payroll WHERE SENDIND = 'N'" +
                                " and period = '" + BusinessLanguage.Period.Trim() +
                                "' and section = 'MANUAL'");

                            panelManualPay.Visible = true;
                            panelLoadFile.Visible = true;

                            cboBonusType.Text = BusinessLanguage.BonusType;
                            lblPeriod2.Text = txtPeriod.Text + " - " + cboSection.Text.Trim();
                            dataGridView1.DataSource = manualPay;

                            int rows = dataGridView1.Rows.Count;
                            if (rows == 1)
                            {
                                btnSendToPayrollDB.Enabled = false;
                                btnManualDelete.Enabled = false;

                            }
                            else
                            {
                                btnSendToPayrollDB.Enabled = true;
                                btnManualDelete.Enabled = true;

                            }


                        }
                        else
                        {
                            //A previous paysend was done, but not send to PalPay yet.
                            extractToBePaysendRecords();

                        }

                    }
                }
                else
                {

                    panelManualPay.Visible = true;
                    panelLoadFile.Visible = true;

                }

            }
        }


        private void hideColumnsOfGrid(string gridname)
        {

            switch (gridname)
            {
                case "dataGridView1":
                    #region payrollgrid
                    if (dataGridView1.Columns.Contains("BUSSUNIT"))
                    {
                        this.dataGridView1.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (dataGridView1.Columns.Contains("MININGTYPE"))
                    {
                        this.dataGridView1.Columns["MININGTYPE"].Visible = false;
                    }
                    if (dataGridView1.Columns.Contains("BONUSTYPE"))
                    {
                        this.dataGridView1.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;
                    #endregion

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) //JVDW
        {
            txtEmployeeNo.Text = dataGridView1["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
            txtEmployeeName.Text = dataGridView1["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
            txtAmount.Text = dataGridView1["EARNINGSVALUE", e.RowIndex].Value.ToString().Trim();

        }

        private void btnManualDelete_Click(object sender, EventArgs e)
        {
            int intRow = dataGridView1.CurrentCell.RowIndex;
            int intColumn = dataGridView1.CurrentCell.ColumnIndex;
            TB.deleteAllExcept(Base.DBConnectionString, "PAYROLL", " WHERE SENDIND = 'N' " +
                                                        "  and period = '" + BusinessLanguage.Period.Trim() + "'");

            TB.saveCalculations2(manualPay, Base.DBConnectionString, "", "PAYROLL");
            string no = dataGridView1["EMPLOYEE_NO", intRow].Value.ToString().Trim();
            //string name = dataGridView1["EMPLOYEE_NAME", intRow].Value.ToString().Trim();

            TB.deleteAllExcept(Base.DBConnectionString, "PAYROLL", "WHERE SENDIND = 'N' AND EMPLOYEE_NO = '" + no.ToString() +
                                                        "' and period = '" + BusinessLanguage.Period.Trim() + "'");


            manualPay = TB.createDataTableWithAdapter(Base.DBConnectionString, "SELECT * FROM PAYROLL WHERE SENDIND = 'N'" +
                                                                               " and period = '" + BusinessLanguage.Period.Trim() + "'");

            //Calculate the batch total
            decimal decBatchtotal = 0;
            foreach (DataRow dr in manualPay.Rows)
            {
                decBatchtotal = decBatchtotal + Convert.ToDecimal(dr["EARNINGSVALUE"].ToString().Trim());
            }
            txtBatchTotal.Text = decBatchtotal.ToString();
            dataGridView1.DataSource = manualPay;
            int rows = dataGridView1.Rows.Count;
            if (rows == 1)
            {
                btnSendToPayrollDB.Enabled = false;
                btnManualDelete.Enabled = false;

            }
            else
            {
                btnSendToPayrollDB.Enabled = true;
                btnManualDelete.Enabled = true;

            }



        }

        private void btnOpenFile_Click(object sender, EventArgs e)//JVDW
        {
            string strFileName = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
            dialog.InitialDirectory = "C:\\iCalc\\Harmony\\Phakisa\\Production\\Data\\ManualPay\\"; dialog.Title = "Select a text file";



            //Present to the user

            if (dialog.ShowDialog() == DialogResult.OK)

                strFileName = dialog.FileName;
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);
            excel.OpenFileShow(strFileName, "");

            if (strFileName == String.Empty)

                return;//user didn't select a file


        }

        private void btnExcelLoad_Click(object sender, EventArgs e)
        {
            manualFileImport();
        }

        public void manualFileImport()
        {

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Phakisa\\Production\\Data\\ManualPay\\ManualPay.xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Phakisa\\Production\\Data\\ManualPay\\ManualPay.xls";

            bool XLSX_exists = File.Exists(FilePath_XLSX);
            bool XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Production\\Data\\ManualPay\\ManualPay.xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Production\\Data\\ManualPay\\ManualPay.xlsx";
            }

            #region extract the sheet name

            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Payfile
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            manualPay.Clear();
            manualPay.Columns.Clear();
            da.Fill(manualPay);


            int count = 0;
            foreach (DataRow col in manualPay.Rows)
            {
                string testArray = col.ItemArray[0].ToString();

                if (testArray == "")
                {
                    manualPay.Rows[count].Delete();
                }
                count++;

            }
            manualPay.AcceptChanges();

            #region remove invalid records

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[manualPay.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }
            //Calculate the batch total
            decimal decBatchtotal = 0;
            foreach (DataRow dr in manualPay.Rows)
            {
                decBatchtotal = decBatchtotal + Convert.ToDecimal(dr["EARNINGSVALUE"].ToString().Trim());
            }

            txtBatchTotal.Text = decBatchtotal.ToString();
            #endregion
            #endregion
            this.Cursor = Cursors.Arrow;



            dataGridView1.DataSource = manualPay;
            int rows = dataGridView1.Rows.Count;
            if (rows == 1)
            {
                btnSendToPayrollDB.Enabled = false;
                btnManualDelete.Enabled = false;

            }
            else
            {
                btnSendToPayrollDB.Enabled = true;
                btnManualDelete.Enabled = true;

            }

        }

        public String[] GetExcelSheetNames(string excelFile)
        {
            //MessageBox.Show(excelFile);
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

                    //MessageBox.Show(row["TABLE_NAME"].ToString());
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
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
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

        private void btnLoad_Click(object sender, EventArgs e)
        {
            string ConCatChar = string.Empty;
            string dateTimeNow = DateTime.Now.ToString();
            string strSign = string.Empty;

            if (txtEmployeeNo.Text.Trim().Length > 0 &&
                txtAmount.Text.Trim().Length > 0)
            {
                if (TB.isStringNumeric(txtAmount.Text.Trim()))
                {
                    #region build the new data
                    DateTime today = DateTime.Today;
                    if (newPayroll.Rows.Count > 0)
                    {

                    }
                    else
                    {
                        newPayroll = TB.createDataTableWithAdapterSelectAll(tstConn.ConnectionString, "Payroll", " where section = 'xxx'");

                        extractReferenceNo();


                        if (TB.RefNo == "new")
                        {
                            TB.RefNo = "001";
                            cboReferenceNo.Text = TB.RefNo.Trim();

                        }
                    }
                    if (Convert.ToDecimal(txtAmount.Text.Trim()) > 0)
                    {
                        strSign = "+";
                    }
                    else
                    {
                        strSign = "-";
                    }

                    string valueEarning = txtAmount.Text.Trim();
                    int comaPlace = valueEarning.LastIndexOf(".");
                    string sub = valueEarning.Substring(comaPlace + 1).ToString();
                    int placesAfterComma = valueEarning.Substring(comaPlace + 1).ToString().Length;


                    if (placesAfterComma != 2 || comaPlace < 0)
                    {
                        if (placesAfterComma == 1)
                        {
                            valueEarning = valueEarning + "0";
                        }
                        else
                        {
                            valueEarning = valueEarning + ".00";
                        }
                    }

                    DataRow newRowPayroll;
                    newRowPayroll = newPayroll.NewRow();

                    newRowPayroll["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                    newRowPayroll["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                    newRowPayroll["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                    newRowPayroll["SECTION"] = "MANUAL";
                    newRowPayroll["GANG"] = "MANUAL PAYSEND";
                    newRowPayroll["PERIOD"] = BusinessLanguage.Period.Trim();
                    newRowPayroll["REFERENCENO"] = TB.RefNo.Trim();
                    newRowPayroll["EARNINGSNAME"] = cboBonusName.Text.Trim();
                    newRowPayroll["TABLENAME"] = cboTableName.Text.Trim();
                    newRowPayroll["EMPLOYEE_NO"] = txtEmployeeNo.Text.Trim();
                    newRowPayroll["EARNINGSCODE"] = cboEarningsCode.Text.Substring(0, 2).Trim();
                    newRowPayroll["PAYMETHODIND"] = cboPaymentInd.Text.Trim().Substring(0, cboPaymentInd.Text.Trim().IndexOf("-"));
                    newRowPayroll["RUNDATE"] = today.ToShortDateString();
                    newRowPayroll["EARNINGSVALUE"] = valueEarning;
                    newRowPayroll["EARNINGSSIGN"] = strSign;
                    newRowPayroll["SENDDATE"] = dateTimeNow.ToString();
                    newRowPayroll["SENDIND"] = "N";
                    newRowPayroll["BATCHNO"] = "Not Send";

                    try
                    {
                        newPayroll.Rows.Add(newRowPayroll);
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message, "insert error", MessageBoxButtons.OK);
                    }

                    newPayroll.AcceptChanges();

                    //Clear - AND to make sure that OUTPUTTABLE is empty.
                    outputTable = new DataTable();
                    outputTable = newPayroll.Copy();

                    this.Cursor = Cursors.Arrow;

                    dataGridView1.DataSource = newPayroll;
                    int rows = dataGridView1.Rows.Count;
                    if (rows == 1)
                    {
                        btnSendToPayrollDB.Enabled = false;
                        btnManualDelete.Enabled = false;
                    }
                    else
                    {
                        btnSendToPayrollDB.Enabled = true;
                        btnManualDelete.Enabled = true;
                    }

                    //Calculate the batch total
                    decimal decBatchtotal = 0;
                    foreach (DataRow dr in outputTable.Rows)
                    {
                        decBatchtotal = decBatchtotal + Convert.ToDecimal(dr["EARNINGSVALUE"].ToString().Trim());
                    }

                    txtBatchTotal.Text = decBatchtotal.ToString();

                    #endregion
                }
            }
        }



        private void btnPrintPayroll_Click(object sender, EventArgs e)
        {
            #region Delete the invalid columns
            DataTable Refrekords = new DataTable();

            if (outputTable.Columns.Contains("TRANSACTIONTYPE"))
            {
                outputTable.Columns.Remove("TRANSACTIONTYPE");
            }
            if (outputTable.Columns.Contains("PAYMENTADVICETYPE"))
            {
                outputTable.Columns.Remove("PAYMENTADVICETYPE");
            }
            if (outputTable.Columns.Contains("PAYMENTUPDATETYPE"))
            {
                outputTable.Columns.Remove("PAYMENTUPDATETYPE");
            }
            if (outputTable.Columns.Contains("PAYMENTIND"))
            {
                outputTable.Columns.Remove("PAYMENTIND");
            }
            if (outputTable.Columns.Contains("BONUSTYPE"))
            {
                outputTable.Columns.Remove("BONUSTYPE");
            }
            if (outputTable.Columns.Contains("PERIOD"))
            {
                outputTable.Columns.Remove("PERIOD");
            }
            if (outputTable.Columns.Contains("SECTION"))
            {
                outputTable.Columns.Remove("SECTION");
            }
            if (outputTable.Columns.Contains("BUSSUNIT"))
            {
                outputTable.Columns.Remove("BUSSUNIT");
            }
            if (outputTable.Columns.Contains("MININGTYPE"))
            {
                outputTable.Columns.Remove("MININGTYPE");
            }
            if (outputTable.Columns.Contains("REFERENCENO"))
            {
            }
            else
            {
                outputTable.Columns.Add("REFERENCENO");
                outputTable.AcceptChanges();
                if (cboReferenceNo.Text.Trim() == "new")
                {
                    cboReferenceNo.Text = "001";
                    foreach (DataRow row in outputTable.Rows)
                    {
                        row["REFERENCENO"] = "001";
                    }
                }
                else
                {
                    foreach (DataRow row in outputTable.Rows)
                    {
                        row["REFERENCENO"] = cboReferenceNo.Text.Trim();
                    }
                }
            }

            outputTable.AcceptChanges();

            #endregion

            #region Add summary column

            lstNames = TB.loadDistinctValuesFromColumn(outputTable, "REFERENCENO");

            if (outputTable.Columns.Contains("SENDIND"))
            {
                try
                {
                    foreach (string s in lstNames)
                    {
                        IEnumerable<DataRow> query2 = from locks in outputTable.AsEnumerable()
                                                      where locks.Field<string>("REFERENCENO").TrimEnd() == s.Trim()
                                                      select locks;

                        Refrekords = query2.CopyToDataTable<DataRow>();

                        decimal totalBonusValue = Convert.ToDecimal(Refrekords.Compute("Sum(BonusValue)", ""));
                        decimal totalEarningsValue = Convert.ToDecimal(Refrekords.Compute("Sum(EarningsValue)", ""));
                        DataRow dr = outputTable.NewRow();

                        dr["EARNINGSVALUE"] = totalEarningsValue.ToString();
                        dr["BONUSVALUE"] = totalBonusValue.ToString();
                        dr["EMPLOYEE_NO"] = "TOTAL";
                        dr["REFERENCENO"] = s;

                        outputTable.Rows.Add(dr);
                        outputTable.AcceptChanges();
                    }
                }
                catch (Exception ex)
                {
                    if (outputTable.Rows[0]["SENDIND"].ToString() == "N")
                    {
                        MessageBox.Show("Please paysend first to get a better print.", "Information", MessageBoxButtons.OK);
                    }
                    else
                    {

                        //MessageBox.Show("Print not possible.  Try again later.", "Information", MessageBoxButtons.OK);
                    }
                }
            }
            else
            {
                if (outputTable.Columns.Contains("BONUSVALUE") && outputTable.Columns.Contains("EARNINGSVALUE"))
                {
                    foreach (string s in lstNames)
                    {
                        IEnumerable<DataRow> query2 = from locks in outputTable.AsEnumerable()
                                                      where locks.Field<string>("REFERENCENO").TrimEnd() == s.Trim()
                                                      select locks;

                        Refrekords = query2.CopyToDataTable<DataRow>();

                        decimal totalBonusValue = Convert.ToDecimal(Refrekords.Compute("Sum(BonusValue)", ""));
                        decimal totalPayrollValue = Convert.ToDecimal(Refrekords.Compute("Sum(PAYROLLVALUE)", ""));
                        decimal totalEarningsValue = Convert.ToDecimal(Refrekords.Compute("Sum(EarningsValue)", ""));
                        DataRow dr = outputTable.NewRow();

                        dr["PAYROLLVALUE"] = totalPayrollValue.ToString();
                        dr["BONUSVALUE"] = totalBonusValue.ToString();
                        dr["EARNINGSVALUE"] = totalEarningsValue.ToString();
                        dr["REFERENCENO"] = s;
                        dr["EMPLOYEE_NO"] = "TOTAL";

                        outputTable.Rows.Add(dr);
                        outputTable.AcceptChanges();
                    }
                }
                else
                {
                    if (outputTable.Columns.Contains("BONUSVALUE"))
                    {
                        foreach (string s in lstNames)
                        {
                            IEnumerable<DataRow> query2 = from locks in outputTable.AsEnumerable()
                                                          where locks.Field<string>("REFERENCENO").TrimEnd() == s.Trim()
                                                          select locks;

                            Refrekords = query2.CopyToDataTable<DataRow>();

                            decimal totalBonusValue = Convert.ToDecimal(Refrekords.Compute("Sum(BonusValue)", ""));
                            decimal totalPayrollValue = Convert.ToDecimal(Refrekords.Compute("Sum(PAYROLLVALUE)", ""));
                            DataRow dr = outputTable.NewRow();

                            dr["PAYROLLVALUE"] = totalPayrollValue.ToString();
                            dr["BONUSVALUE"] = totalBonusValue.ToString();
                            dr["REFERENCENO"] = s;
                            dr["EMPLOYEE_NO"] = "TOTAL";

                            outputTable.Rows.Add(dr);
                            outputTable.AcceptChanges();
                        }
                    }
                    else
                    {
                        if (outputTable.Columns.Contains("EARNINGSVALUE"))
                        {
                            try
                            {
                                foreach (string s in lstNames)
                                {
                                    IEnumerable<DataRow> query2 = from locks in outputTable.AsEnumerable()
                                                                  where locks.Field<string>("REFERENCENO").TrimEnd() == s.Trim()
                                                                  select locks;

                                    Refrekords = query2.CopyToDataTable<DataRow>();

                                    decimal totalBonusValue = Convert.ToDecimal(Refrekords.Compute("Sum(EARNINGSVALUE)", ""));
                                    DataRow dr = outputTable.NewRow();

                                    dr["EARNINGSVALUE"] = totalBonusValue.ToString();
                                    dr["REFERENCENO"] = s;
                                    dr["EMPLOYEE_NO"] = "TOTAL";

                                    outputTable.Rows.Add(dr);
                                    outputTable.AcceptChanges();
                                }
                            }
                            catch
                            {
                                foreach (string s in lstNames)
                                {
                                    IEnumerable<DataRow> query2 = from locks in outputTable.AsEnumerable()
                                                                  where locks.Field<string>("REFERENCENO").TrimEnd() == s.Trim()
                                                                  select locks;

                                    Refrekords = query2.CopyToDataTable<DataRow>();

                                    decimal totalEarningsValue = Convert.ToDecimal(Refrekords.Compute("Sum(EarningsValue)", ""));
                                    DataRow dr = outputTable.NewRow();

                                    dr["EARNINGSVALUE"] = totalEarningsValue.ToString();
                                    dr["REFERENCENO"] = s;
                                    dr["EMPLOYEE_NO"] = "TOTAL";
                                    outputTable.Rows.Add(dr);
                                    outputTable.AcceptChanges();
                                }
                            }
                        }

                    }
                }
            }

            #endregion

            if (outputTable.Rows.Count > 0)
            {
                printHTML(outputTable, "Paysend Summary");

                outputTable.Rows.Clear();
                //cboSection_SelectedIndexChanged("Method", null);
                btnGenerate_Click("Method", null);

            }
            else
            {
                MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
            }
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
                            //FrontDecorator(HTMLWriter);

                            HTMLWriter.WriteLine("Phakisa - " + TabName + "  For : " + cboPayrollGroup.Text.Trim() + "  --- " + cboEarningsColumnName.Text.Trim());
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("=====================================================================");
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("MiningType: " + BusinessLanguage.MiningType + "  --------     BonusType: " + BusinessLanguage.BonusType);
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("Period    : " + BusinessLanguage.Period + "  --------------     Section:  " + cboSection.Text.Trim());
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("Date Printed    : " + DateTime.Today.ToLongDateString().ToString().Trim() + " - " + DateTime.Now.ToShortTimeString().ToString().Trim());
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();

                            grid.RenderControl(HTMLWriter);
                            //RearDecorator(HTMLWriter);

                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();


                    System.Diagnostics.Process P = new System.Diagnostics.Process();
                    P.StartInfo.WorkingDirectory = "C:\\Program Files\\Internet Explorer";
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

    }
}



