using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Phakisa
{
    public partial class scrQuerySQL : Form
    {
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        SqlConnection tstConn = new SqlConnection();
        string tstConnString = string.Empty;

        public scrQuerySQL()
        {
            InitializeComponent();
        }

        internal void TestSQL(SqlConnection myConn, clsGeneral.clsGeneral Gen, string ConnString)
        {
            General = Gen;
            richTextBox1.Text = Gen.textTestSQL;
            tstConn = myConn;
            tstConnString = ConnString;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = General.testSQL(tstConnString, richTextBox1.Text.Trim());
        }

    }
}

