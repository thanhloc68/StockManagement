namespace StockManagerment
{
    partial class StockManagerment
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.lbTongList = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtSearchName = new System.Windows.Forms.TextBox();
            this.cbbSheetStock = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dgvListDb = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.txtposition = new System.Windows.Forms.TextBox();
            this.btnAddShelt = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.cbbSheet = new System.Windows.Forms.ComboBox();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.txtduongdan = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnImportDB = new System.Windows.Forms.Button();
            this.btnopen = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListDb)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnBack);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1618, 634);
            this.panel1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.lbTongList);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtSearchName);
            this.groupBox2.Controls.Add(this.cbbSheetStock);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.dgvListDb);
            this.groupBox2.Location = new System.Drawing.Point(807, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(808, 628);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Danh Sách Sản Phẩm";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(708, 140);
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
            this.label5.Location = new System.Drawing.Point(47, 87);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 13);
            this.label5.TabIndex = 19;
            this.label5.Text = "Tìm Kiếm";
            // 
            // txtSearchName
            // 
            this.txtSearchName.Location = new System.Drawing.Point(103, 51);
            this.txtSearchName.Multiline = true;
            this.txtSearchName.Name = "txtSearchName";
            this.txtSearchName.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSearchName.Size = new System.Drawing.Size(439, 95);
            this.txtSearchName.TabIndex = 18;
            this.txtSearchName.TextChanged += new System.EventHandler(this.txtSearchName_TextChanged);
            // 
            // cbbSheetStock
            // 
            this.cbbSheetStock.FormattingEnabled = true;
            this.cbbSheetStock.Location = new System.Drawing.Point(103, 20);
            this.cbbSheetStock.Name = "cbbSheetStock";
            this.cbbSheetStock.Size = new System.Drawing.Size(214, 21);
            this.cbbSheetStock.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "Lọc Theo Sheet";
            // 
            // dgvListDb
            // 
            this.dgvListDb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListDb.Location = new System.Drawing.Point(6, 169);
            this.dgvListDb.Name = "dgvListDb";
            this.dgvListDb.RowHeadersWidth = 60;
            this.dgvListDb.Size = new System.Drawing.Size(796, 454);
            this.dgvListDb.TabIndex = 9;
            this.dgvListDb.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvListDb_CellFormatting);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.txtposition);
            this.groupBox1.Controls.Add(this.btnAddShelt);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cbbSheet);
            this.groupBox1.Controls.Add(this.dgvData);
            this.groupBox1.Controls.Add(this.txtduongdan);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnImportDB);
            this.groupBox1.Controls.Add(this.btnopen);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(3, 33);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(782, 598);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "UpLoad Excel";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(69, 75);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(201, 23);
            this.textBox1.TabIndex = 23;
            // 
            // txtposition
            // 
            this.txtposition.Location = new System.Drawing.Point(588, 20);
            this.txtposition.Multiline = true;
            this.txtposition.Name = "txtposition";
            this.txtposition.Size = new System.Drawing.Size(73, 23);
            this.txtposition.TabIndex = 21;
            // 
            // btnAddShelt
            // 
            this.btnAddShelt.Location = new System.Drawing.Point(667, 20);
            this.btnAddShelt.Name = "btnAddShelt";
            this.btnAddShelt.Size = new System.Drawing.Size(81, 23);
            this.btnAddShelt.TabIndex = 22;
            this.btnAddShelt.Text = "Thêm Vị Trí";
            this.btnAddShelt.UseVisualStyleBackColor = true;
            this.btnAddShelt.Click += new System.EventHandler(this.btnAddShelt_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 79);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 13);
            this.label4.TabIndex = 19;
            this.label4.Text = "Vị Trí Kệ";
            // 
            // cbbSheet
            // 
            this.cbbSheet.FormattingEnabled = true;
            this.cbbSheet.Location = new System.Drawing.Point(69, 48);
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
            this.txtduongdan.Location = new System.Drawing.Point(68, 19);
            this.txtduongdan.Name = "txtduongdan";
            this.txtduongdan.Size = new System.Drawing.Size(403, 20);
            this.txtduongdan.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Sheet";
            // 
            // btnImportDB
            // 
            this.btnImportDB.Location = new System.Drawing.Point(276, 76);
            this.btnImportDB.Name = "btnImportDB";
            this.btnImportDB.Size = new System.Drawing.Size(94, 23);
            this.btnImportDB.TabIndex = 9;
            this.btnImportDB.Text = "Import to DB";
            this.btnImportDB.UseVisualStyleBackColor = true;
            this.btnImportDB.Click += new System.EventHandler(this.btnImportDB_Click);
            // 
            // btnopen
            // 
            this.btnopen.Location = new System.Drawing.Point(477, 19);
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
            this.label1.Location = new System.Drawing.Point(9, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "File Name";
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(15, 4);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 12;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // StockManagerment
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1642, 650);
            this.Controls.Add(this.panel1);
            this.Name = "StockManagerment";
            this.Text = "StockManagerment";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.StockManagerment_FormClosed);
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListDb)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgvListDb;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cbbSheet;
        private System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.TextBox txtduongdan;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnImportDB;
        private System.Windows.Forms.Button btnopen;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbbSheetStock;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtposition;
        private System.Windows.Forms.Button btnAddShelt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtSearchName;
        private System.Windows.Forms.Label lbTongList;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnBack;
    }
}