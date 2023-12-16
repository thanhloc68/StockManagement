using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StockManagerment
{
    public partial class TikTokInfo : Form
    {
        StockDataContext dbcontext = new StockDataContext();
        public TikTokInfo()
        {
            InitializeComponent();
            LoadDbList();
        }
        public void LoadDbList()
        {
            var list = dbcontext.tbTikTokInfos.Take(50).ToList();
            dgvListDbTikTok.DataSource = list;
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

        private void cbbSheettiktok_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableCollection[cbbSheettiktok.SelectedItem.ToString()];
                dgvDataTiktok.DataSource = dt;
            }
            catch (Exception)
            {
            }
        }

        private void btnImportDBTiktok_Click(object sender, EventArgs e)
        {
            List<tbTikTokInfo> tbTikTokInfos = new List<tbTikTokInfo>();
            string productid, productName, skuid, variationValue, sellerSku;
            int Price, Quantity;
            for (int i = 0; i < dgvDataTiktok.Rows.Count - 1; i++)
            {
                productid = dgvDataTiktok.Rows[i].Cells[0].Value.ToString();
                productName = dgvDataTiktok.Rows[i].Cells[1].Value.ToString();
                skuid = dgvDataTiktok.Rows[i].Cells[2].Value.ToString();
                variationValue = dgvDataTiktok.Rows[i].Cells[3].Value.ToString();
                Price = Convert.ToInt32(dgvDataTiktok.Rows[i].Cells[4].Value.ToString());
                Quantity = Convert.ToInt32(dgvDataTiktok.Rows[i].Cells[5].Value.ToString());
                sellerSku = dgvDataTiktok.Rows[i].Cells[6].Value.ToString();
                var listed = dbcontext.tbTikTokInfos.Any(x => x.seller_sku == sellerSku);
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

                dbcontext.tbTikTokInfos.InsertOnSubmit(st);
                dbcontext.SubmitChanges();
            }
            MessageBox.Show("Đã cập nhật dữ liệu xong", "Thông Báo", MessageBoxButtons.OK);
            LoadDbList();
        }

        private void btnUpdateTiktok_Click(object sender, EventArgs e)
        {

        }

        private async void txtSearchNameTiktok_TextChanged(object sender, EventArgs e)
        {
            await Task.Delay(300);
            dgvListDbTikTok.DataSource = null;
            string textsearch = txtSearchNameTiktok.Text;
            string[] delimeter = { Environment.NewLine };
            string[] findmultitext = textsearch.Split(delimeter, StringSplitOptions.None);
            List<tbTikTokInfo> listproductInStocks = new List<tbTikTokInfo>();
            for (int i = 0; i < findmultitext.Length; i++)
            {
                //var listSearch = from p in dbcontext.tbShopeeInfos where p.SKUProduct.Contains(findmultitext[i]) select p;
                var listSearch = dbcontext.tbTikTokInfos.Where(x => x.seller_sku.Contains(findmultitext[i])).ToList();
                foreach (var item in listSearch)
                {
                    item.quantity = 0;
                }
                listproductInStocks.AddRange(listSearch);
            }
            dgvListDbTikTok.DataSource = listproductInStocks;
        }

        private void btnSoldOutTiktok_Click(object sender, EventArgs e)
        {

        }

        private void btnExportTiktok_Click(object sender, EventArgs e)
        {
            ExportList();
        }
        public async void ExportList()
        {
            await Task.Delay(500);
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
            worksheet.Cells[1, 1] = "product_id"; worksheet.Cells[1, 2] = "product_name"; worksheet.Cells[1, 3] = "sku_id"; worksheet.Cells[1, 4] = "variation_value"; worksheet.Cells[1, 5] = "price"; worksheet.Cells[1, 6] = "quantity"; worksheet.Cells[1, 7] = "seller_sku";
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
            workbook.SaveAs("d:\\Project\\xuatharavanvashopee\\updateShopee", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            // Exit from the application  
            app.Quit();
        }
    }
}
