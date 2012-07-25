using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Base = clsMain;

namespace Phakisa
{
    public partial class scrSecurity : Form
    {
        DataTable userid = new DataTable();
        DataTable usrAccess = new DataTable();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsTable.clsTable TB = new clsTable.clsTable();
        SqlConnection BaseConn = new SqlConnection();
        string DBConnectionString = string.Empty;
        string BaseConnectionString = string.Empty;
        string AnalysisConnectionString = string.Empty;
        string strUserid = string.Empty;
        string strEnvironment = string.Empty;
        List<string> lstNames = new List<string>();

        public scrSecurity()
        {
            InitializeComponent();
        }

        internal void userAccessLoad(SqlConnection Conn,clsMain.clsMain Main,clsTable.clsTable clsTable,string usr,string Environment)
        {
            listBox1.Enabled = false;
            listBox2.Enabled = false;
            listBox3.Enabled = false;
            BaseConn = Conn;
            Base = Main;
            TB = clsTable;
            strUserid = usr;
            strEnvironment = Environment;
          

            DataTable userids = Base.SelectUserAccess(Base.BaseConnectionString);

            lstNames = TB.loadDistinctValuesFromColumn(userids, "USERID");

            foreach (string s in lstNames)
            {
                if (s != strUserid)
                {
                    comboBox1.Items.Add(s.ToString().Trim());
                }
            }

            //Fill listbox LEVEL1 with distict values
            DataTable level1 = Base.SelectUserAccess(Base.BaseConnectionString);

            lstNames = TB.loadDistinctValuesFromColumn(userids, "LEVEL1");

            foreach (string s in lstNames)
            {
                listBox1.Items.Add(s.ToString().Trim());
            }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox3.Items.Clear();

            DataTable level2 = Base.GetLevel2(listBox1.SelectedItem.ToString().Trim(), Base.BaseConnectionString);

            foreach (DataRow r in level2.Rows)
            {
                if (listBox2.Items.Contains(r["LEVEL2"].ToString().Trim()))
                {
                }
                else
                {
                    listBox2.Items.Add(r["LEVEL2"].ToString().Trim());
                }
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();

            DataTable level3 = Base.GetDataBy11(listBox1.SelectedItem.ToString().Trim(), listBox2.SelectedItem.ToString().Trim(), Base.BaseConnectionString);

            foreach (DataRow r in level3.Rows)
            {
                if (listBox3.Items.Contains(r["LEVEL3"].ToString().Trim() + " - " + r["CODENAME"].ToString().Trim()))
                {
                }
                else
                {
                    listBox3.Items.Add(r["LEVEL3"].ToString().Trim() + " - " + r["CODENAME"].ToString().Trim());
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Enabled = false;
            listBox2.Enabled = false;
            listBox3.Enabled = false;
            comboBox2.Enabled = true;
            comboBox3.Enabled = false;
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();

            usrAccess = Base.SelectAccessByUserid(comboBox1.Text.Trim(), Base.BaseConnectionString);
            dataGridView1.DataSource = usrAccess;

            if (usrAccess.Rows.Count > 0)
            {
                comboBox2.Items.Add(usrAccess.Rows[0]["BussUnit"].ToString().Trim());
                comboBox3.Items.Add(usrAccess.Rows[0]["Resp"].ToString().Trim());
            }
            else
            {
                ////Load all possible bussinessunits
                //DataTable dt = Base.SelectBussUnit(Base.BaseConnectionString);
                //foreach (DataRow r in dt.Rows)
                //{
                //    comboBox2.Items.Add(r["BUSSUNIT"].ToString().Trim() + " - " + r["BUSSUNIT_DESC"].ToString().Trim());
                //}

                ////Load all possible responsibilities
                //DataTable dt2 = Base.SelectResponsibilities(Base.BaseConnectionString);
                //foreach (DataRow rr in dt2.Rows)
                //{

                //    comboBox3.Items.Add(rr["RESP"].ToString().Trim() + " - " + rr["RESP_DESC"].ToString().Trim());

                //}
            }
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //Delete all for userid

            Base.DeleteAllforUserid(comboBox1.Text.Trim(), Base.BaseConnectionString);

            //Insert all on the datagrid back into the database.

            foreach (DataRow r in usrAccess.Rows)
            {
                try
                {

                    Base.InsertUsrAccess(strEnvironment, r["USERID"].ToString().Trim(), r["RESP"].ToString().Trim(),comboBox2.Text.Substring(0,2),
                        r["LEVEL1"].ToString().Trim(), r["LEVEL2"].ToString().Trim(), r["LEVEL3"].ToString().Trim(), 
                        r["CODENAME"].ToString().Trim(), "Y", Base.BaseConnectionString);
                   
                }
                catch
                {
                    MessageBox.Show(r["USERID"].ToString().Trim() + " " + r["RESP"].ToString().Trim() + " " + 
                                    r["BUSSUNIT"].ToString().Trim() + " " + r["LEVEL1"].ToString().Trim() + " " + 
                                    r["LEVEL2"].ToString().Trim() + " " + r["LEVEL3"].ToString().Trim() + " " + 
                                    r["CODENAME"].ToString().Trim() + " " + "Y could not be inserted.", 
                                    "Access insert Error!!!", MessageBoxButtons.OK);
                }

            }
             
            MessageBox.Show("Access successfully granted.", "Confirmation", MessageBoxButtons.OK);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //remove record from userid datatable
            int intRow = dataGridView1.CurrentCell.RowIndex;

            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1[i, intRow].Style.BackColor = Color.Lavender;
            }

            DialogResult result = MessageBox.Show("Continue deleting the access shown?", "Confirm", MessageBoxButtons.OKCancel);

            switch (result)
            {
                case DialogResult.OK:

                    string strSQL = "Delete from usrAccess where Bussunit = '" + dataGridView1["BUSSUNIT",intRow].Value.ToString().Trim().Substring(0,2)+
                                    "' and Resp = '" + dataGridView1["RESP", intRow].Value.ToString().Trim() +
                                    "' and Userid = '" + dataGridView1["USERID", intRow].Value.ToString().Trim() +
                                    "' and Level1 = '" + dataGridView1["LEVEL1", intRow].Value.ToString().Trim() +
                                    "' and Level2 = '" + dataGridView1["LEVEL2", intRow].Value.ToString().Trim() +
                                    "' and Level3 = '" + dataGridView1["LEVEL3", intRow].Value.ToString().Trim() +
                                    "' and Codename = '" + dataGridView1["CODENAME", intRow].Value.ToString().Trim() +
                                    "' and Ind = '" + dataGridView1["IND", intRow].Value.ToString().Trim() + "'";

                    Base.InsertData(Base.BaseConnectionString,strSQL); 
                    usrAccess.Rows.Remove(usrAccess.Rows[intRow]);
                    usrAccess.AcceptChanges();
                    dataGridView1.DataSource = usrAccess;
                    break;

                case DialogResult.Cancel:
                    break;

            }

            if (dataGridView1.SelectedRows.Count == 1)
            {
                int introw = dataGridView1.Rows.IndexOf(dataGridView1.SelectedRows[0]);

                usrAccess.Rows.Remove(usrAccess.Rows[introw]);
                usrAccess.AcceptChanges();
                dataGridView1.DataSource = usrAccess;
            }
           
        }

        private void btnGrantAccess_Click(object sender, EventArgs e)
        {
            DataRow dr = usrAccess.NewRow();

            string[] strError = { "", "", "" };
            DataColumn[] keys = new DataColumn[8];
            keys[0] =  usrAccess.Columns[0] ;
            keys[1] =  usrAccess.Columns[1];
            keys[2] =  usrAccess.Columns[2];
            keys[3] = usrAccess.Columns[3];
            keys[4] = usrAccess.Columns[4];
            keys[5] = usrAccess.Columns[5];
            keys[6] = usrAccess.Columns[6];
                
            usrAccess.PrimaryKey = keys;   ///check

            dr[0] = comboBox1.Text.Trim();
            dr[1] = comboBox3.Text.Trim();
            dr[2] = comboBox2.Text.Trim().Substring(0, comboBox2.Text.Trim().IndexOf('-')); 
            dr[3] = listBox1.SelectedItem.ToString().Trim();
            dr[4] = listBox2.SelectedItem.ToString().Trim();
            dr[5] = listBox3.SelectedItem.ToString().Trim().Substring(0, listBox3.SelectedItem.ToString().Trim().IndexOf('-') - 1);
            dr[6] = listBox3.SelectedItem.ToString().Trim().Substring(listBox3.SelectedItem.ToString().Trim().IndexOf('-') + 2);
            dr[7] = 'Y';

            //comboBox1.Text.Trim(),comboBox3.Text.Trim().Substring(0, comboBox3.Text.Trim().IndexOf('-') - 1), comboBox2.Text.Trim().Substring(0, comboBox2.Text.Trim().IndexOf('-') - 1), listBox1.SelectedItem.ToString().Trim(),listBox2.SelectedItem.ToString().Trim(),listBox3.SelectedItem.ToString().Trim().Substring(0, listBox3.SelectedItem.ToString().Trim().IndexOf('-') - 1),listBox3.SelectedItem.ToString().Trim().Substring(listBox3.SelectedItem.ToString().Trim().IndexOf('-') + 2), 'Y';


            try
            {
                usrAccess.Rows.Add(dr.ItemArray);
                usrAccess.AcceptChanges();
            }
            catch
            {

                MessageBox.Show("User has already access to the current selection.", "Duplication Error", MessageBoxButtons.OK);
            }

            dataGridView1.DataSource = usrAccess;
          
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            comboBox3.Enabled = true;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true; 
        }


        }
    }

