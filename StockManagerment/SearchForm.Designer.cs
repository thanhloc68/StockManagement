namespace StockManagerment
{
    partial class SearchForm
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
            this.cbbSheet = new System.Windows.Forms.ComboBox();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.txtduongdan = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnImportDB = new System.Windows.Forms.Button();
            this.btnopen = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.lbTongList = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtSearchName = new System.Windows.Forms.TextBox();
            this.dgvListDb = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnSoldOut = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnback = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListDb)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbbSheet
            // 
            this.cbbSheet.FormattingEnabled = true;
            this.cbbSheet.Location = new System.Drawing.Point(69, 76);
            this.cbbSheet.Name = "cbbSheet";
            this.cbbSheet.Size = new System.Drawing.Size(201, 21);
            this.cbbSheet.TabIndex = 14;
            this.cbbSheet.SelectedIndexChanged += new System.EventHandler(this.cbbSheet_SelectedIndexChanged);
            // 
            // dgvData
            // 
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Location = new System.Drawing.Point(6, 105);
            this.dgvData.Name = "dgvData";
            this.dgvData.Size = new System.Drawing.Size(770, 518);
            this.dgvData.TabIndex = 8;
            this.dgvData.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvData_CellFormatting);
            // 
            // txtduongdan
            // 
            this.txtduongdan.Location = new System.Drawing.Point(68, 47);
            this.txtduongdan.Name = "txtduongdan";
            this.txtduongdan.Size = new System.Drawing.Size(403, 20);
            this.txtduongdan.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Sheet";
            // 
            // btnImportDB
            // 
            this.btnImportDB.Location = new System.Drawing.Point(287, 76);
            this.btnImportDB.Name = "btnImportDB";
            this.btnImportDB.Size = new System.Drawing.Size(94, 23);
            this.btnImportDB.TabIndex = 9;
            this.btnImportDB.Text = "Import to DB";
            this.btnImportDB.UseVisualStyleBackColor = true;
            this.btnImportDB.Click += new System.EventHandler(this.btnImportDB_Click);
            // 
            // btnopen
            // 
            this.btnopen.Location = new System.Drawing.Point(477, 47);
            this.btnopen.Name = "btnopen";
            this.btnopen.Size = new System.Drawing.Size(54, 23);
            this.btnopen.TabIndex = 11;
            this.btnopen.Text = "...";
            this.btnopen.UseVisualStyleBackColor = true;
            this.btnopen.Click += new System.EventHandler(this.btnopen_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "File Name";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(708, 16);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(94, 23);
            this.btnExport.TabIndex = 23;
            this.btnExport.Text = "Export Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // lbTongList
            // 
            this.lbTongList.AutoSize = true;
            this.lbTongList.Location = new System.Drawing.Point(716, 150);
            this.lbTongList.Name = "lbTongList";
            this.lbTongList.Size = new System.Drawing.Size(0, 13);
            this.lbTongList.TabIndex = 20;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(27, 57);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 13);
            this.label5.TabIndex = 19;
            this.label5.Text = "Tìm Kiếm";
            // 
            // txtSearchName
            // 
            this.txtSearchName.Location = new System.Drawing.Point(83, 21);
            this.txtSearchName.Multiline = true;
            this.txtSearchName.Name = "txtSearchName";
            this.txtSearchName.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSearchName.Size = new System.Drawing.Size(439, 71);
            this.txtSearchName.TabIndex = 18;
            this.txtSearchName.TextChanged += new System.EventHandler(this.txtSearchName_TextChanged);
            // 
            // dgvListDb
            // 
            this.dgvListDb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListDb.Location = new System.Drawing.Point(6, 105);
            this.dgvListDb.Name = "dgvListDb";
            this.dgvListDb.RowHeadersWidth = 60;
            this.dgvListDb.Size = new System.Drawing.Size(796, 517);
            this.dgvListDb.TabIndex = 9;
            this.dgvListDb.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvListDb_CellFormatting);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnSoldOut);
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.lbTongList);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtSearchName);
            this.groupBox2.Controls.Add(this.dgvListDb);
            this.groupBox2.Location = new System.Drawing.Point(807, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(808, 628);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Danh Sách Sản Phẩm";
            // 
            // btnSoldOut
            // 
            this.btnSoldOut.Location = new System.Drawing.Point(708, 76);
            this.btnSoldOut.Name = "btnSoldOut";
            this.btnSoldOut.Size = new System.Drawing.Size(94, 23);
            this.btnSoldOut.TabIndex = 24;
            this.btnSoldOut.Text = "Hạ Tồn";
            this.btnSoldOut.UseVisualStyleBackColor = true;
            this.btnSoldOut.Click += new System.EventHandler(this.btnSoldOut_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnUpdate);
            this.groupBox1.Controls.Add(this.cbbSheet);
            this.groupBox1.Controls.Add(this.dgvData);
            this.groupBox1.Controls.Add(this.txtduongdan);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnImportDB);
            this.groupBox1.Controls.Add(this.btnopen);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(3, 41);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(782, 590);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "UpLoad Excel";
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(387, 76);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(94, 23);
            this.btnUpdate.TabIndex = 15;
            this.btnUpdate.Text = "Update DB";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnback);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1618, 634);
            this.panel1.TabIndex = 1;
            // 
            // btnback
            // 
            this.btnback.Location = new System.Drawing.Point(9, 3);
            this.btnback.Name = "btnback";
            this.btnback.Size = new System.Drawing.Size(54, 23);
            this.btnback.TabIndex = 16;
            this.btnback.Text = "Back";
            this.btnback.UseVisualStyleBackColor = true;
            this.btnback.Click += new System.EventHandler(this.btnback_Click);
            // 
            // SearchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1638, 655);
            this.Controls.Add(this.panel1);
            this.Name = "SearchForm";
            this.Text = "SearchForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SearchForm_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListDb)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ComboBox cbbSheet;
        private System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.TextBox txtduongdan;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnImportDB;
        private System.Windows.Forms.Button btnopen;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label lbTongList;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtSearchName;
        private System.Windows.Forms.DataGridView dgvListDb;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnSoldOut;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnback;
    }
}