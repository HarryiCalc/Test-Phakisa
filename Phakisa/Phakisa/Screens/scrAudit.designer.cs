namespace Phakisa
{
    partial class scrAudit
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnExtract = new System.Windows.Forms.Button();
            this.cboType = new System.Windows.Forms.ComboBox();
            this.cboTableName = new System.Windows.Forms.ComboBox();
            this.grdAuditSheet = new System.Windows.Forms.DataGridView();
            this.btnAuditReport = new System.Windows.Forms.Button();
            this.panel4 = new System.Windows.Forms.Panel();
            this.txtMiningType = new System.Windows.Forms.TextBox();
            this.txtBonusType = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblMintype = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDatabaseName = new System.Windows.Forms.TextBox();
            this.txtUserDetails = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.grdAuditSheet)).BeginInit();
            this.panel4.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExtract
            // 
            this.btnExtract.Location = new System.Drawing.Point(793, 9);
            this.btnExtract.Name = "btnExtract";
            this.btnExtract.Size = new System.Drawing.Size(121, 35);
            this.btnExtract.TabIndex = 0;
            this.btnExtract.Text = "Extract";
            this.btnExtract.UseVisualStyleBackColor = true;
            this.btnExtract.Click += new System.EventHandler(this.btnExtract_Click);
            // 
            // cboType
            // 
            this.cboType.FormattingEnabled = true;
            this.cboType.Items.AddRange(new object[] {
            "D - Delete",
            "I - Insert",
            "U - Update",
            "A - All Types"});
            this.cboType.Location = new System.Drawing.Point(77, 8);
            this.cboType.Name = "cboType";
            this.cboType.Size = new System.Drawing.Size(121, 21);
            this.cboType.TabIndex = 5;
            this.cboType.Text = "U - Update";
            // 
            // cboTableName
            // 
            this.cboTableName.FormattingEnabled = true;
            this.cboTableName.Location = new System.Drawing.Point(77, 36);
            this.cboTableName.Name = "cboTableName";
            this.cboTableName.Size = new System.Drawing.Size(121, 21);
            this.cboTableName.TabIndex = 10;
            // 
            // grdAuditSheet
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.grdAuditSheet.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.grdAuditSheet.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.grdAuditSheet.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.grdAuditSheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdAuditSheet.Location = new System.Drawing.Point(28, 93);
            this.grdAuditSheet.Name = "grdAuditSheet";
            this.grdAuditSheet.Size = new System.Drawing.Size(1177, 369);
            this.grdAuditSheet.TabIndex = 12;
            this.grdAuditSheet.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.grdAuditSheet_CellMouseClick);
            // 
            // btnAuditReport
            // 
            this.btnAuditReport.Location = new System.Drawing.Point(920, 9);
            this.btnAuditReport.Name = "btnAuditReport";
            this.btnAuditReport.Size = new System.Drawing.Size(121, 35);
            this.btnAuditReport.TabIndex = 13;
            this.btnAuditReport.Text = "Report";
            this.btnAuditReport.UseVisualStyleBackColor = true;
            this.btnAuditReport.Click += new System.EventHandler(this.btnAuditReport_Click);
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.AliceBlue;
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.txtMiningType);
            this.panel4.Controls.Add(this.txtBonusType);
            this.panel4.Controls.Add(this.label3);
            this.panel4.Controls.Add(this.lblMintype);
            this.panel4.Location = new System.Drawing.Point(367, 8);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(207, 69);
            this.panel4.TabIndex = 33;
            // 
            // txtMiningType
            // 
            this.txtMiningType.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtMiningType.Enabled = false;
            this.txtMiningType.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMiningType.Location = new System.Drawing.Point(77, 9);
            this.txtMiningType.Name = "txtMiningType";
            this.txtMiningType.Size = new System.Drawing.Size(115, 20);
            this.txtMiningType.TabIndex = 32;
            // 
            // txtBonusType
            // 
            this.txtBonusType.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtBonusType.Enabled = false;
            this.txtBonusType.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBonusType.Location = new System.Drawing.Point(77, 39);
            this.txtBonusType.Name = "txtBonusType";
            this.txtBonusType.Size = new System.Drawing.Size(115, 20);
            this.txtBonusType.TabIndex = 31;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Bonus Type:";
            // 
            // lblMintype
            // 
            this.lblMintype.AutoSize = true;
            this.lblMintype.Location = new System.Drawing.Point(4, 9);
            this.lblMintype.Name = "lblMintype";
            this.lblMintype.Size = new System.Drawing.Size(68, 13);
            this.lblMintype.TabIndex = 10;
            this.lblMintype.Text = "Mining Type:";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.AliceBlue;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtDatabaseName);
            this.panel1.Controls.Add(this.txtUserDetails);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Location = new System.Drawing.Point(28, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(333, 69);
            this.panel1.TabIndex = 32;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "User Details :";
            // 
            // txtDatabaseName
            // 
            this.txtDatabaseName.Enabled = false;
            this.txtDatabaseName.Location = new System.Drawing.Point(104, 41);
            this.txtDatabaseName.Name = "txtDatabaseName";
            this.txtDatabaseName.Size = new System.Drawing.Size(224, 20);
            this.txtDatabaseName.TabIndex = 3;
            // 
            // txtUserDetails
            // 
            this.txtUserDetails.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtUserDetails.Enabled = false;
            this.txtUserDetails.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUserDetails.Location = new System.Drawing.Point(104, 11);
            this.txtUserDetails.Name = "txtUserDetails";
            this.txtUserDetails.Size = new System.Drawing.Size(224, 20);
            this.txtUserDetails.TabIndex = 29;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 43);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(90, 13);
            this.label9.TabIndex = 10;
            this.label9.Text = "Database Name :";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.AliceBlue;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.cboType);
            this.panel2.Controls.Add(this.cboTableName);
            this.panel2.Location = new System.Drawing.Point(580, 8);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(207, 69);
            this.panel2.TabIndex = 34;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "TableName:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 9);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(34, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Type:";
            // 
            // scrAudit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1232, 497);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnAuditReport);
            this.Controls.Add(this.grdAuditSheet);
            this.Controls.Add(this.btnExtract);
            this.Name = "scrAudit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Audit Reports";
            ((System.ComponentModel.ISupportInitialize)(this.grdAuditSheet)).EndInit();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExtract;
        private System.Windows.Forms.ComboBox cboType;
        private System.Windows.Forms.ComboBox cboTableName;
        private System.Windows.Forms.DataGridView grdAuditSheet;
        private System.Windows.Forms.Button btnAuditReport;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.TextBox txtMiningType;
        private System.Windows.Forms.TextBox txtBonusType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblMintype;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDatabaseName;
        private System.Windows.Forms.TextBox txtUserDetails;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
    }
}

