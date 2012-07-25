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
    public partial class scrProfile : Form
    {
        clsBL.clsBL bl = new clsBL.clsBL();
        clsMain.clsMain Base = new clsMain.clsMain();

        SqlConnection BaseConn = new SqlConnection();

        public scrProfile()
        {
            InitializeComponent();
        }

        internal void FormLoad(clsBL.clsBL Business, SqlConnection Conn)
        {
            bl = Business;
            BaseConn = Conn;
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Add a User
            if (textBox1.Text.Trim().Length > 0)
            {
                object intCount = Base.CountUserid(textBox1.Text.Trim().ToUpper(), BaseConn.ConnectionString);

                if ((int)intCount > 0)
                {
                    DialogResult result = MessageBox.Show("Userid already exists.  Do you want to do an update on the password?", "Confirm", MessageBoxButtons.YesNo);

                    switch (result)
                    {
                        case DialogResult.Yes:
                            if (textBox2.Text.Trim().Length == 0)
                            {
                                MessageBox.Show("Please supply a password.", "Information", MessageBoxButtons.OK);
                            }
                            else
                            {

                                string strmessage = bl.validatePassword(textBox2.Text.Trim());
                                if (strmessage.Trim().Length > 0)
                                {
                                    MessageBox.Show(strmessage, "Password update failed", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    DateTime today = DateTime.Today;
                                    try
                                    {
                                        Base.UpdatePassword(textBox2.Text.Trim(), today.ToShortDateString(), bl.Userid.Trim(), BaseConn.ConnectionString);
                                        MessageBox.Show("Password succesfully updated", "Confirmation", MessageBoxButtons.OK);
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Password updated failed", "Password Error", MessageBoxButtons.OK);
                                    }
                                }
                            }
                            break;

                        case DialogResult.No:
                            break;
                    }
                }

                else
                {
                    DialogResult result = MessageBox.Show("Add a new user?", "Confirm", MessageBoxButtons.YesNo);

                    switch (result)
                    {
                        case DialogResult.Yes:

                            string strmessage = bl.validatePassword(textBox2.Text.Trim());
                            if (strmessage.Trim().Length > 0)
                            {
                                MessageBox.Show(strmessage, "Password validataion failed", MessageBoxButtons.OK);
                            }
                            else
                            {
                                DateTime today = DateTime.Today;
                                try
                                {
                                    Base.InsertProfile(textBox1.Text.Trim().ToUpper(), textBox2.Text.Trim(), today.ToShortDateString(), BaseConn.ConnectionString);
                                    MessageBox.Show("User successfully added.", "Confirmation", MessageBoxButtons.OK);
                                }
                                catch
                                {
                                    MessageBox.Show("ADD function failed.", "User Add Error", MessageBoxButtons.OK);
                                }
                            }
                            break;

                        case DialogResult.No:
                            break;


                    }
                }
            }
            else
            {
                MessageBox.Show("Enter a userid.", "Information", MessageBoxButtons.OK);
            }
        }
    }
}
