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
    public partial class SearchForm : Form
    {
        StockDataContext dbcontext = new StockDataContext();
        public SearchForm()
        {
            InitializeComponent();
            LoadDbList();
        }
        public void LoadDbList()
        {
            try
            {
                var list = dbcontext.tbShopeeInfos.ToList();
                dgvListDb.DataSource = list;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi dữ liệu tại -> " + ex.Message);
            }
        }
        private void SearchForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
        DataTableCollection tableCollection;
        private void btnopen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtduongdan.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cbbSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                cbbSheet.Items.Add(table.TableName);
                            }
                        }
                    }
                }
            }
        }

        private void cbbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableCollection[cbbSheet.SelectedItem.ToString()];
                dgvData.DataSource = dt;
            }
            catch (Exception)
            {
            }
        }

        private void btnImportDB_Click(object sender, EventArgs e)
        {
            List<tbShopeeInfo> tbShopeeInfos = new List<tbShopeeInfo>();
            string productCode, productName, classificationCode, classificationName, SKUProduct, SKU;
            int Price, Quantity;
            for (int i = 0; i < dgvData.Rows.Count - 1; i++)
            {
                productCode = dgvData.Rows[i].Cells[0].Value.ToString();
                productName = dgvData.Rows[i].Cells[1].Value.ToString();
                classificationCode = dgvData.Rows[i].Cells[2].Value.ToString();
                classificationName = dgvData.Rows[i].Cells[3].Value.ToString();
                SKUProduct = dgvData.Rows[i].Cells[4].Value.ToString();
                SKU = dgvData.Rows[i].Cells[5].Value.ToString();
                Price = Convert.ToInt32(dgvData.Rows[i].Cells[6].Value.ToString());
                Quantity = Convert.ToInt32(dgvData.Rows[i].Cells[7].Value.ToString());
                var listed = dbcontext.tbShopeeInfos.Any(x => x.SKUProduct == SKUProduct);
                var st = new tbShopeeInfo
                {
                    productCode = productCode,
                    productName = productName,
                    classificationCode = classificationCode,
                    classificationName = classificationName,
                    SKUProduct = SKUProduct,
                    SKU = SKU,
                    Price = Price,
                    Quantity = Quantity
                };
                if (listed) continue;
              
                dbcontext.tbShopeeInfos.InsertOnSubmit(st);
                dbcontext.SubmitChanges();
            }
            MessageBox.Show("Đã cập nhật dữ liệu xong", "Thông Báo", MessageBoxButtons.OK);
            LoadDbList();
        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {

        }
        private void btnExport_Click(object sender, EventArgs e)
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
            worksheet.Cells[1, 1] = "et_title_product_id"; worksheet.Cells[1, 2] = "et_title_product_name"; worksheet.Cells[1, 3] = "et_title_variation_id"; worksheet.Cells[1, 4] = "et_title_variation_name"; worksheet.Cells[1, 5] = "et_title_parent_sku"; worksheet.Cells[1, 6] = "et_title_variation_sku"; worksheet.Cells[1, 7] = "et_title_variation_price"; worksheet.Cells[1, 8] = "et_title_variation_stock"; worksheet.Cells[1, 9] = "et_title_reason";
            worksheet.Cells[2, 1] = "sales_info"; worksheet.Cells[2, 2] = "4d93e627870723759fffa6927c542c0e"; worksheet.Cells[2, 3] = "0";
            // storing header part in Excel
            for (int i = 1; i < dgvListDb.Columns.Count ; i++)
            {
                worksheet.Cells[3, i] = dgvListDb.Columns[i].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgvListDb.Rows.Count; i++)
            {
                for (int j = 1; j < dgvListDb.Columns.Count; j++)
                {
                    worksheet.Cells[i + 7, j] = dgvListDb.Rows[i].Cells[j].Value?.ToString();
                }
            }
            worksheet.Range["A:A"].NumberFormat = 0;
            worksheet.Range["C:C"].NumberFormat = 0;
            worksheet.Range["E:E"].NumberFormat = 0;
            // save the application  
            app.AskToUpdateLinks = false;
            app.DisplayAlerts = false;
            workbook.SaveAs("d:\\xuatharavanvashopee\\updateShopee", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            // Exit from the application  
            app.Quit();
        }

        private void dgvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvData.Rows[e.RowIndex].HeaderCell.Value = Convert.ToString(e.RowIndex + 1);
        }

        private async void txtSearchName_TextChanged(object sender, EventArgs e)
        {
            await Task.Delay(300);
            dgvListDb.DataSource = null;
            string textsearch = txtSearchName.Text;
            string[] delimeter = { Environment.NewLine };
            string[] findmultitext = textsearch.Split(delimeter, StringSplitOptions.None);
            List<tbShopeeInfo> listproductInStocks = new List<tbShopeeInfo>();
            for (int i = 0; i < findmultitext.Length; i++)
            {
                //var listSearch = from p in dbcontext.tbShopeeInfos where p.SKUProduct.Contains(findmultitext[i]) select p;
                var listSearch = dbcontext.tbShopeeInfos.Where(x=> x.SKUProduct.Contains(findmultitext[i])).ToList();
                foreach (var item in listSearch)
                {
                    item.Quantity = 0;
                }
                listproductInStocks.AddRange(listSearch);
            }
            dgvListDb.DataSource = listproductInStocks;
        }

        private void dgvListDb_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvListDb.Rows[e.RowIndex].HeaderCell.Value = Convert.ToString(e.RowIndex + 1);
        }

        private void btnback_Click(object sender, EventArgs e)
        {
            Main mains = new Main();
            mains.Show();
            this.Hide();
        }

        private void btnSoldOut_Click(object sender, EventArgs e)
        {

        }
    }
}
