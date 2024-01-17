﻿using ExcelDataReader;
using StockManagerment.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace StockManagerment
{
    public partial class TikTokInfo : Form
    {
        private readonly Stock_ManagementEntities _dbContext = new Stock_ManagementEntities();
        public TikTokInfo(Stock_ManagementEntities dbContext)
        {
            _dbContext = dbContext;
        }
        public TikTokInfo()
        {
            InitializeComponent();
            LoadDbList();
        }
        public async void LoadDbList()
        {
            try
            {
                var list = await _dbContext.tbTikTokInfoes.Take(50).ToListAsync();
                dgvListDbTikTok.DataSource = list;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }

        }
        private void btnbackForm_Click(object sender, EventArgs e)
        {
            Main mains = new Main();
            mains.Show();
            this.Hide();
        }
        DataTableCollection tableCollection;
        private void btnopentiktok_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtduongdantiktok.Text = openFileDialog.FileName;
                        using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                });
                                tableCollection = result.Tables;
                                cbbSheettiktok.Items.Clear();
                                foreach (DataTable table in tableCollection)
                                {
                                    cbbSheettiktok.Items.Add(table.TableName);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void cbbSheettiktok_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableCollection[cbbSheettiktok.SelectedItem.ToString()];
                DataTable dtNew = dt.Clone();
                for (int i = 4; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    dtNew.Rows.Add(row.ItemArray);
                }
                dgvDataTiktok.DataSource = dtNew;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private async void btnImportDBTiktok_Click(object sender, EventArgs e)
        {
            try
            {
                List<tbTikTokInfo> tbTikTokInfos = new List<tbTikTokInfo>();
                string productid, category, productName, skuid, variationValue, sellerSku;
                int Price, Quantity;
                for (int i = 0; i < dgvDataTiktok.Rows.Count - 1; i++)
                {
                    productid = dgvDataTiktok.Rows[i].Cells[0].Value.ToString();
                    productName = dgvDataTiktok.Rows[i].Cells[1].Value.ToString();
                    category = dgvDataTiktok.Rows[i].Cells[2].Value.ToString();
                    skuid = dgvDataTiktok.Rows[i].Cells[3].Value.ToString();
                    variationValue = dgvDataTiktok.Rows[i].Cells[4].Value.ToString();
                    Price = Convert.ToInt32(dgvDataTiktok.Rows[i].Cells[5].Value.ToString());
                    Quantity = Convert.ToInt32(dgvDataTiktok.Rows[i].Cells[6].Value.ToString());
                    sellerSku = dgvDataTiktok.Rows[i].Cells[7].Value.ToString();
                    var listed = _dbContext.tbTikTokInfoes.Any(x => x.seller_sku == sellerSku);
                    var st = new tbTikTokInfo
                    {
                        product_id = productid,
                        product_name = productName,
                        sku_id = skuid,
                        variation_value = variationValue,
                        price = Price,
                        quantity = Quantity,
                        seller_sku = sellerSku,
                    };
                    if (listed) continue;
                    _dbContext.tbTikTokInfoes.AddOrUpdate(st);
                    await _dbContext.SaveChangesAsync();
                }
                MessageBox.Show("Đã cập nhật dữ liệu xong", "Thông Báo", MessageBoxButtons.OK);
                LoadDbList();
            }
            catch (Exception ex) { messageData(ex); }
        }
        private void btnUpdateTiktok_Click(object sender, EventArgs e)
        {

        }
        private async void txtSearchNameTiktok_TextChanged(object sender, EventArgs e)
        {
            try
            {
                await Task.Delay(300);
                dgvListDbTikTok.DataSource = null;
                string textsearch = txtSearchNameTiktok.Text;
                string[] delimeter = { Environment.NewLine };
                string[] findmultitext = textsearch.Split(delimeter, StringSplitOptions.None);
                List<tbTikTokInfo> listproductInStocks = new List<tbTikTokInfo>();
                var productsTiktok = await _dbContext.tbTikTokInfoes.ToListAsync();

                foreach (var term in findmultitext)
                {
                    //var listSearch = from p in dbcontext.tbShopeeInfos where p.SKUProduct.Contains(findmultitext[i]) select p;
                    var filteredProducts = productsTiktok.Where(x => x.seller_sku.Contains(term)).ToList();
                    foreach (var item in filteredProducts)
                    {
                        item.quantity = 0;
                    }
                    listproductInStocks.AddRange(filteredProducts);
                }
                dgvListDbTikTok.DataSource = listproductInStocks;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void btnSoldOutTiktok_Click(object sender, EventArgs e)
        {

        }
        private void btnExportTiktok_Click(object sender, EventArgs e)
        {
            try
            {
                ExportList();
            }
            catch (Exception ex) { messageData(ex); }
        }
        public async void ExportList()
        {
            try
            {
                await Task.Delay(400);
                // creating Excel Application  
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program  
                app.Visible = true;
                app.AlertBeforeOverwriting = false;
                app.DisplayAlerts = false;
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet  
                // get the reference of first sheet. By default its name is Sheet2.  
                worksheet.Name = @"Sheet1";
                worksheet.Cells[1, 1] = "product_id"; worksheet.Cells[1, 2] = "category"; worksheet.Cells[1, 3] = "product_name"; worksheet.Cells[1, 4] = "sku_id"; worksheet.Cells[1, 5] = "variation_value"; worksheet.Cells[1, 6] = "price"; worksheet.Cells[1, 7] = "quantity"; worksheet.Cells[1, 8] = "seller_sku";
                worksheet.Cells[2, 1] = "V3"; worksheet.Cells[2, 2] = "Sales_Information";
                // storing header part in Excel
                worksheet.Range["A:A"].NumberFormat = "@";
                worksheet.Range["C:C"].NumberFormat = "@";
                for (int i = 1; i < dgvListDbTikTok.Columns.Count; i++)
                {
                    worksheet.Cells[3, i] = dgvListDbTikTok.Columns[i].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                for (int i = 0; i < dgvListDbTikTok.Rows.Count; i++)
                {
                    for (int j = 1; j < dgvListDbTikTok.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 6, j] = dgvListDbTikTok.Rows[i].Cells[j].Value?.ToString();
                    }
                }
                // save the application  
                app.AskToUpdateLinks = false;
                app.DisplayAlerts = false;
                workbook.SaveAs("d:\\Project\\xuatharavanvashopee\\updateTikTok", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                // Exit from the application  
                app.Quit();
            }
            catch (Exception ex) { messageData(ex); }
        }
        private void messageData(Exception ex)
        {
            MessageBox.Show($"Lỗi dữ liệu {ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
