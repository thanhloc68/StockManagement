namespace StockManagerment
{
    partial class Main
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
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.btnStockManagerment = new System.Windows.Forms.Button();
            this.btninsertupdateShopee = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Location = new System.Drawing.Point(55, 42);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(150, 25);
            this.btnExportExcel.TabIndex = 1;
            this.btnExportExcel.Text = "Thêm và Xuất file Excel Shopee & Haravan";
            this.btnExportExcel.UseVisualStyleBackColor = true;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // btnStockManagerment
            // 
            this.btnStockManagerment.Location = new System.Drawing.Point(211, 42);
            this.btnStockManagerment.Name = "btnStockManagerment";
            this.btnStockManagerment.Size = new System.Drawing.Size(150, 25);
            this.btnStockManagerment.TabIndex = 2;
            this.btnStockManagerment.Text = "Quản lý Kho 23";
            this.btnStockManagerment.UseVisualStyleBackColor = true;
            this.btnStockManagerment.Click += new System.EventHandler(this.btnStockManagerment_Click);
            // 
            // btninsertupdateShopee
            // 
            this.btninsertupdateShopee.Location = new System.Drawing.Point(367, 42);
            this.btninsertupdateShopee.Name = "btninsertupdateShopee";
            this.btninsertupdateShopee.Size = new System.Drawing.Size(150, 25);
            this.btninsertupdateShopee.TabIndex = 3;
            this.btninsertupdateShopee.Text = "Cập nhật shopee";
            this.btninsertupdateShopee.UseVisualStyleBackColor = true;
            this.btninsertupdateShopee.Click += new System.EventHandler(this.btninsertupdateShopee_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btninsertupdateShopee);
            this.Controls.Add(this.btnStockManagerment);
            this.Controls.Add(this.btnExportExcel);
            this.Name = "Main";
            this.Text = "Main";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.Button btnStockManagerment;
        private System.Windows.Forms.Button btninsertupdateShopee;
    }
}

