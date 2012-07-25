namespace Phakisa
{
    partial class scrLogon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(scrLogon));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnProcess = new System.Windows.Forms.Button();
            this.cboMiningType = new System.Windows.Forms.ComboBox();
            this.lblMininType = new System.Windows.Forms.Label();
            this.cboBonusType = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboPeriods = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboUserid = new System.Windows.Forms.ComboBox();
            this.cboBussUnit = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPassw = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblVersion = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.cboEnvironment = new System.Windows.Forms.ComboBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.txtRegion = new System.Windows.Forms.TextBox();
            this.panelSignon = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panelSignon.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.cboMiningType);
            this.panel1.Controls.Add(this.lblMininType);
            this.panel1.Controls.Add(this.cboBonusType);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.cboPeriods);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Enabled = false;
            this.panel1.Location = new System.Drawing.Point(3, 354);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(626, 101);
            this.panel1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.btnProcess);
            this.panel3.Location = new System.Drawing.Point(510, 53);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(109, 37);
            this.panel3.TabIndex = 53;
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(16, 6);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(75, 23);
            this.btnProcess.TabIndex = 0;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // cboMiningType
            // 
            this.cboMiningType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMiningType.FormattingEnabled = true;
            this.cboMiningType.Location = new System.Drawing.Point(84, 10);
            this.cboMiningType.Name = "cboMiningType";
            this.cboMiningType.Size = new System.Drawing.Size(154, 21);
            this.cboMiningType.TabIndex = 15;
            this.cboMiningType.SelectedIndexChanged += new System.EventHandler(this.cboMiningType_SelectedIndexChanged);
            // 
            // lblMininType
            // 
            this.lblMininType.AutoSize = true;
            this.lblMininType.Location = new System.Drawing.Point(10, 13);
            this.lblMininType.Name = "lblMininType";
            this.lblMininType.Size = new System.Drawing.Size(68, 13);
            this.lblMininType.TabIndex = 14;
            this.lblMininType.Text = "Mining Type:";
            // 
            // cboBonusType
            // 
            this.cboBonusType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboBonusType.FormattingEnabled = true;
            this.cboBonusType.Location = new System.Drawing.Point(325, 10);
            this.cboBonusType.Name = "cboBonusType";
            this.cboBonusType.Size = new System.Drawing.Size(154, 21);
            this.cboBonusType.TabIndex = 13;
            this.cboBonusType.SelectedIndexChanged += new System.EventHandler(this.cboBonusType_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(253, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Bonus Type:";
            // 
            // cboPeriods
            // 
            this.cboPeriods.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPeriods.FormattingEnabled = true;
            this.cboPeriods.Location = new System.Drawing.Point(84, 47);
            this.cboPeriods.Name = "cboPeriods";
            this.cboPeriods.Size = new System.Drawing.Size(154, 21);
            this.cboPeriods.TabIndex = 11;
            this.cboPeriods.SelectedIndexChanged += new System.EventHandler(this.cboPeriods_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Period :";
            // 
            // cboUserid
            // 
            this.cboUserid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboUserid.FormattingEnabled = true;
            this.cboUserid.Location = new System.Drawing.Point(63, 40);
            this.cboUserid.Name = "cboUserid";
            this.cboUserid.Size = new System.Drawing.Size(154, 21);
            this.cboUserid.TabIndex = 9;
            this.cboUserid.SelectedIndexChanged += new System.EventHandler(this.cboUserid_SelectedIndexChanged);
            // 
            // cboBussUnit
            // 
            this.cboBussUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboBussUnit.FormattingEnabled = true;
            this.cboBussUnit.Location = new System.Drawing.Point(325, 7);
            this.cboBussUnit.Name = "cboBussUnit";
            this.cboBussUnit.Size = new System.Drawing.Size(154, 21);
            this.cboBussUnit.TabIndex = 8;
            this.cboBussUnit.SelectedIndexChanged += new System.EventHandler(this.cboBussUnit_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(238, 7);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(82, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Bussiness Unit :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 7);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "Region :";
            // 
            // txtPassw
            // 
            this.txtPassw.BackColor = System.Drawing.Color.Lavender;
            this.txtPassw.Location = new System.Drawing.Point(325, 40);
            this.txtPassw.Name = "txtPassw";
            this.txtPassw.PasswordChar = '*';
            this.txtPassw.Size = new System.Drawing.Size(154, 20);
            this.txtPassw.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(240, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Password :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Userid :";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(16, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Signon";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.lblVersion);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Controls.Add(this.panel5);
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Location = new System.Drawing.Point(12, 12);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(645, 460);
            this.panel2.TabIndex = 5;
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.BackColor = System.Drawing.Color.White;
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.Location = new System.Drawing.Point(477, 167);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(149, 13);
            this.lblVersion.TabIndex = 55;
            this.lblVersion.Text = "Development Version 1.2";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.BackColor = System.Drawing.Color.White;
            this.label9.Font = new System.Drawing.Font("Verdana", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(467, 31);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(162, 32);
            this.label9.TabIndex = 55;
            this.label9.Text = "Phakisa";
            // 
            // panel5
            // 
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel5.Controls.Add(this.label10);
            this.panel5.Controls.Add(this.label8);
            this.panel5.Controls.Add(this.cboEnvironment);
            this.panel5.Location = new System.Drawing.Point(3, 195);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(626, 53);
            this.panel5.TabIndex = 54;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(249, 16);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(142, 18);
            this.label10.TabIndex = 53;
            this.label10.Text = "Environment = ";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(10, 18);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 13);
            this.label8.TabIndex = 7;
            this.label8.Text = "Env :";
            // 
            // cboEnvironment
            // 
            this.cboEnvironment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboEnvironment.FormattingEnabled = true;
            this.cboEnvironment.Items.AddRange(new object[] {
            "Production",
            "Test",
            "Development"});
            this.cboEnvironment.Location = new System.Drawing.Point(63, 15);
            this.cboEnvironment.Name = "cboEnvironment";
            this.cboEnvironment.Size = new System.Drawing.Size(154, 21);
            this.cboEnvironment.TabIndex = 8;
            this.cboEnvironment.SelectedIndexChanged += new System.EventHandler(this.cboEnvironment_SelectedIndexChanged);
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.txtRegion);
            this.panel4.Controls.Add(this.cboBussUnit);
            this.panel4.Controls.Add(this.panelSignon);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Controls.Add(this.label6);
            this.panel4.Controls.Add(this.label5);
            this.panel4.Controls.Add(this.txtPassw);
            this.panel4.Controls.Add(this.label2);
            this.panel4.Controls.Add(this.cboUserid);
            this.panel4.Location = new System.Drawing.Point(3, 254);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(626, 94);
            this.panel4.TabIndex = 53;
            // 
            // txtRegion
            // 
            this.txtRegion.BackColor = System.Drawing.Color.Lavender;
            this.txtRegion.Enabled = false;
            this.txtRegion.Location = new System.Drawing.Point(63, 4);
            this.txtRegion.Name = "txtRegion";
            this.txtRegion.Size = new System.Drawing.Size(154, 20);
            this.txtRegion.TabIndex = 53;
            this.txtRegion.Text = "FREEGOLD";
            // 
            // panelSignon
            // 
            this.panelSignon.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panelSignon.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelSignon.Controls.Add(this.button1);
            this.panelSignon.Location = new System.Drawing.Point(510, 43);
            this.panelSignon.Name = "panelSignon";
            this.panelSignon.Size = new System.Drawing.Size(109, 37);
            this.panelSignon.TabIndex = 52;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(3, 16);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(626, 173);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 21;
            this.pictureBox1.TabStop = false;
            // 
            // scrLogon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(672, 484);
            this.Controls.Add(this.panel2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "scrLogon";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Logon";
            this.Load += new System.EventHandler(this.scrLogon_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.scrLogon_FormClosing);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panelSignon.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtPassw;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboBussUnit;
        private System.Windows.Forms.ComboBox cboUserid;
        private System.Windows.Forms.ComboBox cboPeriods;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ComboBox cboMiningType;
        private System.Windows.Forms.Label lblMininType;
        private System.Windows.Forms.ComboBox cboBonusType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panelSignon;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cboEnvironment;
        private System.Windows.Forms.TextBox txtRegion;
        private System.Windows.Forms.Label lblVersion;
    }
}