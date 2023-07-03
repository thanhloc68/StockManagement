using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Linq.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StockManagerment
{
    public partial class StockManagerment : Form
    {
        StockDataContext dbStock = new StockDataContext();
        public StockManagerment()
        {
            InitializeComponent();
            LoadDbList();
            loadCBB();

        }
        public void LoadDbList()
        {
            var list = dbStock.productInStocks.ToList();
            dgvListDb.DataSource = list;
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

        private void btnImportDB_Click(object sender, EventArgs e)
        {
            string name, sku;
            int quantity;
            int indexShelt;
            for (int i = 0; i < dgvData.Rows.Count-1; i++)
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

                dbStock.productInStocks.InsertOnSubmit(st);
                dbStock.SubmitChanges();
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

        private async void txtSearchName_TextChanged(object sender, EventArgs e)
        {
            await Task.Delay(400);
            string textsearch = txtSearchName.Text;
            string[] delimeter = { Environment.NewLine };
            string[] findmultitext = textsearch.Split(delimeter, StringSplitOptions.None);
            List<productInStock> listproductInStocks = new List<productInStock>();
            for (int i = 0; i < findmultitext.Length; i++)
            {
                var listSearch = from p in dbStock.productInStocks where p.sku.Contains(findmultitext[i]) select p;
                listproductInStocks.AddRange(listSearch);
            }
            dgvListDb.DataSource = listproductInStocks;
        }

        private void dgvListDb_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvListDb.Rows[e.RowIndex].HeaderCell.Value = System.Convert.ToString(e.RowIndex + 1);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportList();
        }
        public void ExportList()
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

        private void dgvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvData.Rows[e.RowIndex].HeaderCell.Value = Convert.ToString(e.RowIndex + 1);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Main mains = new Main();
            mains.Show();
            this.Hide();
        }
    }
}