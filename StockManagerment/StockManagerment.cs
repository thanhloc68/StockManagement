using ExcelDataReader;
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
    public partial class StockManagerment : Form
    {
        private readonly Stock_ManagementEntities _dbContext = new Stock_ManagementEntities();
        public StockManagerment(Stock_ManagementEntities dbContext)
        {
            _dbContext = dbContext;
        }
        public StockManagerment()
        {
            InitializeComponent();
            LoadDbList();
            loadCBB();
        }
        public async void LoadDbList()
        {
            try
            {
                var list = await _dbContext.productInStocks.ToListAsync();
                dgvListDb.DataSource = list;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        public void loadCBB()
        {
            //var listSheet = from p in dbStock.Shelts select new { name = p.Name, id = p.id };
            //cbbSheetStock.DataSource = listSheet.OrderBy(x => x.name).ToList();
            //cbbSheetStock.DisplayMember = "Name";
        }
        private void btnUpdate_Click(object sender, EventArgs e)
        {

        }
        private void StockManagerment_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
        private async void btnImportDB_Click(object sender, EventArgs e)
        {
            try
            {

                string name, sku;
                int quantity;
                int indexShelt;
                for (int i = 0; i < dgvData.Rows.Count - 1; i++)
                {
                    sku = dgvData.Rows[i].Cells[1].Value.ToString();
                    name = dgvData.Rows[i].Cells[2].Value.ToString();
                    quantity = Convert.ToInt32(dgvData.Rows[i].Cells[4].Value.ToString());
                    indexShelt = Convert.ToInt32(dgvData.Rows[i].Cells[7].Value.ToString());
                    var st = new productInStock
                    {
                        name_Product = name,
                        sku = sku,
                        Stock = quantity,
                        Shelf = indexShelt
                    };

                    _dbContext.productInStocks.AddOrUpdate(st);
                    await _dbContext.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void btnAddShelt_Click(object sender, EventArgs e)
        {
            //Shelt shelt = new Shelt();
            //var dataSheet = dbStock.Shelts.Where(x => x.Name == x.Name);
            //shelt.Name = txtposition.Text.ToString();
            //foreach (var item in dataSheet)
            //{
            //    if (item.Name == txtposition.Text.ToString()) return;
            //}

            //dbStock.Shelts.InsertOnSubmit(shelt);
            //dbStock.SubmitChanges();
            //loadCBB();
        }
        DataTableCollection tableCollection;
        private void btnopen_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
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
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void cbbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableCollection[cbbSheet.SelectedItem.ToString()];
                dgvData.DataSource = dt;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }

        }
        private async void txtSearchName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                await Task.Delay(400);
                string textsearch = txtSearchName.Text;
                string[] delimeter = { Environment.NewLine };
                string[] findmultitext = textsearch.Split(delimeter, StringSplitOptions.None);
                // Fetch all products into memory
                var products = await _dbContext.productInStocks.ToListAsync();
                List<productInStock> listproductInStocks = new List<productInStock>();
                foreach (var term in findmultitext)
                {
                    // Filter in-memory using LINQ to Objects
                    var listSearch = products.Where(x => x.sku.Contains(term)).ToList();
                    listproductInStocks.AddRange(listSearch);
                }
                dgvListDb.DataSource = listproductInStocks;
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void dgvListDb_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                dgvListDb.Rows[e.RowIndex].HeaderCell.Value = System.Convert.ToString(e.RowIndex + 1);
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                ExportList();
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        public void ExportList()
        {
            try
            {
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

                worksheet.Name = @"Export File";
                /* worksheet.Cells[1, 1] = "100643 - ";*/

                // storing header part in Excel
                for (int i = 1; i < dgvListDb.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dgvListDb.Columns[i - 1].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                for (int i = 0; i < dgvListDb.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dgvListDb.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dgvListDb.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                // save the application  

                app.AskToUpdateLinks = false;
                app.DisplayAlerts = false;
                workbook.SaveAs("d:\\xuatharavanvashopee\\KiemtraKe", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                // Exit from the application  
                app.Quit();
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void dgvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                dgvData.Rows[e.RowIndex].HeaderCell.Value = Convert.ToString(e.RowIndex + 1);
            }
            catch (Exception ex)
            {
                messageData(ex);
            }
        }
        private void btnBack_Click(object sender, EventArgs e)
        {
            Main mains = new Main();
            mains.Show();
            this.Hide();
        }
        public void messageData(Exception ex)
        {
            MessageBox.Show($"Lỗi dữ liệu {ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}