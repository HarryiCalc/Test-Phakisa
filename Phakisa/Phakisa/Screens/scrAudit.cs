using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;


namespace Phakisa
{
    public partial class scrAudit : Form
    {
        public DataTable dt = new DataTable();
        ArrayList words = new ArrayList();
        ArrayList header = new ArrayList();
        ArrayList columnLength = new ArrayList();
        clsMain.clsMain Base = new clsMain.clsMain(); 
        DataTable AuditLog = new DataTable();
        string strHeader = string.Empty;
        string strWords = string.Empty; 
        clsTable.clsTable TB = new clsTable.clsTable();
        SqlConnection myConn = new SqlConnection();

        string DBConnString = string.Empty;
        string BaseConnString = string.Empty;

        clsBL.clsBL BusinessLanguage = new clsBL.clsBL(); 
        DataTable BonusTable = new DataTable();
        DataTable outputTable = new DataTable();
        DataTable Configs = new DataTable();
        DataTable EarningCodes = new DataTable();
        DataTable manualPay = new DataTable();
        DataColumn myDataColumn = new DataColumn();
        List<string> lstNames = new List<string>();
        string strMetaReportCode = "BSFnupmWkNxm8ZAA1ZhlOgL8fNdMdg4zhJj/j6T0vEyG9aSzk/HPwYcrjmawRGou66hBtseT7qJE+9hbEq9jces6bcGJmtz4Ih8Fic4UIw0Kt2lEffc05nFdiD2aQC0m";

        string strSection = string.Empty;
        string[] sectionlist = new string[10];
        string sectioncbo = string.Empty;
        string globalSection = string.Empty; 
        
        public scrAudit()
        {
            InitializeComponent();
        }

        internal void AuditLoad(string ConnString, string BConnString, clsBL.clsBL classBL, clsTable.clsTable classTable,
                                string section)
        {
            string globalSection = section; 
            TB = classTable;
            BusinessLanguage = classBL;

            DBConnString = ConnString;
            BaseConnString = BConnString;
            sectioncbo = section;

            txtUserDetails.Text = BusinessLanguage.Userid + " - " + BusinessLanguage.Region + " - " + BusinessLanguage.BussUnit;
            txtMiningType.Text = BusinessLanguage.MiningType;
            txtBonusType.Text = BusinessLanguage.BonusType;
            txtDatabaseName.Text = TB.DBName;

            lstNames = TB.getAllFromSelectedTable(DBConnString, "AUDIT", "Tablename");

            foreach (string s in lstNames)
            {
                cboTableName.Items.Add(s.Trim());

            }
        }

        private void btnExtract_Click(object sender, EventArgs e)
        {
            DataRow dr;
            string strSQL = string.Empty;

            Base.DBName = TB.DBName;
                   
            //string ConnString = "Server=VAATJIE-PC\\SQLEXPRESS;Trusted_Connection = True;User Id=jvdw;password =go;database=Phakisa201006";
            //Server=.\\ANGLOPLATS;Trusted_Connection = True;User Id=am;password =am;database=Phakisa201104
            //string ConnString = "Server=AMLAPTOP-PC\angloplats;Trusted_Connection = True;User Id=am;password =am;database=" + Base.DBName;
            //string ConnString = "Server=.\\ANGLOPLATS;Trusted_Connection = True;User Id=am;password =am;database=Phakisa201104";

            #region extract auditlog
            SqlConnection myConn = new SqlConnection(DBConnString);
            try
            {
                if (cboType.Text.Substring(0, 1) == "A" && cboTableName.Text.Trim() == "ALL")
                {
                    strSQL = "select * from Audit";
                }
                else
                {
                    if (cboType.Text.Substring(0, 1) == "A")
                    {
                        strSQL = "select * from Audit where tablename = '" + cboTableName.Text.Trim() + "'";
                    }
                    else
                    {
                        if (cboTableName.Text.Trim() == "ALL")
                        {
                            strSQL = "select * from Audit where Type = '" + cboType.Text.Trim().Substring(0, 1) + "' ";                    
                        }
                        else
                        {
                            strSQL = "select * from Audit where Type = '" + cboType.Text.Trim().Substring(0, 1) +
                                     "' and tablename = '" + cboTableName.Text.Trim() + "'";
                        }
                    }
                }
                myConn.Open();

                dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(strSQL, DBConnString);
                da.Fill(dt);
                TB.createAuditLogTable(DBConnString);
                TB.deleteAllExcept(DBConnString, "AUDITLOG"); 
                //Get the lay-out of the AUDIT talbe
                AuditLog = TB.createDataTableWithAdapterSelectAll(DBConnString, "AUDITLOG");
                myConn.Close();

            }
            catch (Exception)
            {

                throw;
            }

            finally { myConn.Close(); }
            #endregion

            if (dt.Rows.Count > 0)
            {

                #region Create HeaderRecord
                //Add the records to the Auditlog table.
                //First the header record
                //The primary key will contain the header record.

                string rowPKIn = dt.Rows[0]["PK"].ToString().Trim();
                //Do while statement
                do
                {
                    rowPKIn = createHeaderArray(rowPKIn);

                } while (rowPKIn.Length > 0);

                //Output the header record to the temp file
                strHeader = string.Empty;

                for (int i = 0; i <= header.Count - 1; i++)
                {
                    strHeader = strHeader + " - " + header[i].ToString().Trim();
                    int inttest  = strHeader.ToString().LastIndexOf("-");
                    columnLength.Add(Convert.ToString(inttest));                
                }

                header.Clear();
                dr = AuditLog.NewRow();
                dr["BUSSUNIT"] = "BUSSUNIT";
                dr["MININGTYPE"] = "MININGTYPE";
                dr["BONUSTYPE"] = "BONUSTYPE";
                dr["SECTION"] = "SECTION";
                dr["PERIOD"] = "PERIOD";
                dr["TABLENAME"] = "TABLENAME";
                dr["TYPE"] = cboType.Text.Trim();
                dr["FIELDNAME"] = "FIELDNAME";
                dr["OLDVALUE"] = "OLDVALUE";
                dr["NEWVALUE"] = "NEWVALUE";
                dr["UPDATEDATE"] = "UPDATEDATE";
                dr["USERNAME"] = "USERNAME";
                dr["RECORDAFFECTED"] = strHeader;
                dr["RECORDTYPE"] = "HEADER";

                AuditLog.Rows.Add(dr);
                AuditLog.AcceptChanges();
                strHeader = string.Empty;

                #endregion

                #region Create DataRecords
                foreach (DataRow rowPK in dt.Rows)
                {
                    words.Clear();
                    rowPKIn = string.Empty;
                    rowPKIn = rowPK["PK"].ToString();

                    

                    //extract the section value
                    do
                    {
                        rowPKIn = extractSection(rowPKIn);

                    } while (rowPKIn.Length > 0);

                    rowPKIn = rowPK["PK"].ToString();

                    //Do while statement
                    do
                    {
                        rowPKIn = createDataArray(rowPKIn);

                    } while (rowPKIn.Length > 0);

                    strWords = string.Empty;

                    //Create the dataArray and space it according to the header record.
                    for (int i = 0; i <= words.Count - 1; i++)
                    {
                        int intlength = strWords.ToString().Length;
                        if (intlength == 0)
                        {
                            strWords = strWords + " - " + words[i].ToString().Trim();
                        }
                        else
                        {
                            if (Convert.ToInt16(intlength) < Convert.ToInt16(columnLength[i].ToString()))
                            {
                                int maxValue = Convert.ToInt16(columnLength[i].ToString()) - Convert.ToInt16(intlength);
                                do
                                {
                                    strWords = strWords + " ";
                                    maxValue = maxValue - 1;
                                } while (maxValue >= 0);
                                

                                strWords = strWords + " - " + words[i].ToString().Trim();
                            }
                            else
                            {
                                strWords = strWords + " - " + words[i].ToString().Trim();
                            }

                        }
                    }

                    string rowType = rowPK["Type"].ToString().Trim();
                    string rowTableName = rowPK["Tablename"].ToString().Trim();
                    string rowFielName = rowPK["Fieldname"].ToString().Trim();
                    string rowOldVal = rowPK["OldValue"].ToString().Trim();
                    string rowNewVal = rowPK["NewValue"].ToString().Trim();
                    string rowUpdateDate = rowPK["UpdateDate"].ToString().Trim();
                    string rowUserName = rowPK["UserName"].ToString().Trim();

                    //DataRow dr;
                    dr = AuditLog.NewRow();
                    dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                    dr["MININGTYPE"] = BusinessLanguage.MiningType;
                    dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                    dr["SECTION"] = strSection;
                    dr["PERIOD"] = BusinessLanguage.Period;
                    dr["TABLENAME"] = rowTableName;
                    dr["TYPE"] = rowType;
                    dr["FIELDNAME"] = rowFielName;
                    dr["OLDVALUE"] = rowOldVal;
                    dr["NEWVALUE"] = rowNewVal;
                    dr["UPDATEDATE"] = rowUpdateDate;
                    dr["USERNAME"] = rowUserName;
                    dr["RECORDAFFECTED"] = strWords;
                    dr["RECORDTYPE"] = "DATA";

                    AuditLog.Rows.Add(dr);
                    AuditLog.AcceptChanges();
                    grdAuditSheet.DataSource = AuditLog;

                }
                #endregion

                TB.saveCalculations2(AuditLog, DBConnString, "", "AUDITLOG");

                grdAuditSheet.DataSource = AuditLog;
            }
            else
            {
                MessageBox.Show("No records extracted for your query.", "Information", MessageBoxButtons.OK);

            }
        }

private void connectToDB()
        {

            if (myConn.State == ConnectionState.Closed)
            {
                try
                {
                    myConn.Open();
                }
                catch(Exception xx)
                {
                    try
                    {
                        int Count = (int)Base.CountDBinMaster(Base.DBName);
                        if (Count == 0)
                        {
                            //Attach the db
                            Base.attachDB(Base.DBName);
                            myConn.Open();

                        }
                        else
                        { }

                    }
                    catch {
                        MessageBox.Show(xx.Message);
                    }
                    MessageBox.Show(xx.Message);
                }
            }

        }

        private string createHeaderArray(string PrimaryKeyString)
        {
            int firts = PrimaryKeyString.IndexOf(">");
            string get = PrimaryKeyString.Substring(0, firts + 1);
            string newStr = get.Replace("<", "");
            string newSrtr2 = newStr.Replace(">", "").Trim();
            int newStr2Lenght = newSrtr2.Length;
            int eqSign = newSrtr2.IndexOf("=");
            string split = newSrtr2.Substring(0, eqSign);
            if (split == "MININGTYPE" || split == "BONUSTYPE" || split == "SECTION" || split == "BUSSUNIT" || split == "PERIOD")
            {

            }
            else
            {
                header.Add(split);
            }
            PrimaryKeyString = PrimaryKeyString.Remove(0, firts + 1);
            return PrimaryKeyString;

        }

        private string createDataArray(string PrimaryKeyString)
        {
            int firts = PrimaryKeyString.IndexOf(">");
            string get = PrimaryKeyString.Substring(0, firts + 1);
            string newStr = get.Replace("<", "");
            string newSrtr2 = newStr.Replace(">", "").Trim();
            int newStr2Lenght = newSrtr2.Length;
            int eqSign = newSrtr2.IndexOf("=");
            string split = newSrtr2.Substring(0, eqSign);
            if (split == "MININGTYPE" || split == "BONUSTYPE" || split == "SECTION" || split == "BUSSUNIT" || split == "PERIOD")
            {

            }
            else
            {
                words.Add(newSrtr2.Remove(0, eqSign + 1));
            }

            PrimaryKeyString = PrimaryKeyString.Remove(0, firts + 1);
            return PrimaryKeyString;

        }

        private string extractSection(string PrimaryKeyString)
        {
            int firts = PrimaryKeyString.IndexOf(">");
            string get = PrimaryKeyString.Substring(0, firts + 1);
            string newStr = get.Replace("<", "");
            string newSrtr2 = newStr.Replace(">", "").Trim();
            int newStr2Lenght = newSrtr2.Length;
            int eqSign = newSrtr2.IndexOf("=");
            string split = newSrtr2.Substring(0, eqSign);
            if (split == "SECTION")
            {
                strSection = newSrtr2.Remove(0, eqSign + 1);

            }
                PrimaryKeyString = PrimaryKeyString.Remove(0, firts + 1);
                return PrimaryKeyString;
            

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

        #endregion

        private void grdAuditSheet_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdAuditSheet);
            }
        }

        private void btnAuditReport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            MetaReportRuntime.AppClass mm = new MetaReportRuntime.AppClass();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\AngloPlats\\Plats\\" + BusinessLanguage.Env + "\\REPORTS\\";
            mm.StartReport("AuditReport");

            this.Cursor = Cursors.Arrow;
        }


    }
    }


