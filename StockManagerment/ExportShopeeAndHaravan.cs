using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Color = System.Drawing.Color;

namespace StockManagerment
{
    public partial class ExportShopeeAndHaravan : Form
    {
        public ExportShopeeAndHaravan()
        {
            InitializeComponent();
            txtLanguage.Text = "Tiếng Việt";
            cbbLoaiBia.Text = "Bìa mềm";
            txtXuatKhau.Text = "Trong nước";
            txtNameProduct.Text = "";
            txtSlShopee.Text = "0";
            txtSlHrv.Text = "0";
            // Read values from file
            var strReadJsonIndustry = File.ReadAllText(@"industry.json");
            // Convert to Json Object
            var getIndustry = JsonConvert.DeserializeObject<List<InDustry>>(strReadJsonIndustry);
            foreach (var value in getIndustry.Select(a => a.Industry_name))
            {
                cbbIndustry.Items.Add(value);
            }
            // Read values from file
            var strReadJson = File.ReadAllText(@"id-brand.json");
            // Convert to Json Object
            var getIdBrand = JsonConvert.DeserializeObject<List<Incident>>(strReadJson);
            foreach (var option in getIdBrand.Select(p => p.name))
            {
                cbbBrand.Items.Add(option);
            }
        }
        void addHaravan(bool hrv)
        {
            if (hrv == false)
            {
                int index = 0;
                DataGridViewRow rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                const string quote = "\"";
                List<string> listCollection = new List<string>();
                if (cbbDanhMuc.Text.ToString() != "")
                {
                    string[] collection = cbbDanhMuc.Text.ToString().Split(',');
                    foreach (var everCollection in collection)
                    {
                        listCollection.Add(everCollection);
                    }
                        listCollection.Add(txtTag.Text.ToString());
                }
                rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString(); rowHaravan.Cells[1].Value = txtNameProduct.Text.ToString();
                rowHaravan.Cells[2].Value = "<h2 style= " + quote + "text-align: justify;" + quote + "><strong>" + txtNameProduct.Text.ToString() + "</strong></h2>" +
                                            "<p style = " + quote + "text-align: justify;" + quote + ">Thông tin sản phẩm </p><table><tbody>" +
                                            "<tr><th style = " + quote + "text-align: justify;" + quote + ">Mã hàng </th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + "> " + txtSKU.Text.ToString() + "</td></tr>" +
                                            "<tr><th style= " + quote + "text-align: justify;" + quote + ">Tên Nhà Cung Cấp</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + "> " + cbbNCC.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">Tác giả</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + cbbBrand.Text.ToString() + "</td>" +
                                            "</tr><tr><th style=" + quote + "text-align: justify;" + quote + ">Người Dịch</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtTrans.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">NXB</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtNPH.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">Năm XB</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtNamsx.Text.ToString() + "</td></tr><tr>" +
                                            "<th style=" + quote + "text-align: justify;" + quote + ">Ngôn Ngữ</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtLanguage.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">Trọng lượng (gr)</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtWeight.Text.ToString() + "</td></tr><tr>" +
                                            "<th style=" + quote + "text-align: justify;" + quote + ">Kích Thước Bao Bì</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + txtSize.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">Số trang</th><td style=" + quote + "text-align:justify" + quote + "> " + txtNumpage.Text.ToString() + "</td></tr>" +
                                            "<tr><th style=" + quote + "text-align: justify;" + quote + ">Hình thức</th>" +
                                            "<td style=" + quote + "text-align:justify" + quote + ">" + cbbLoaiBia.Text.ToString() + "</td></tr></tbody></table>" +
                                            "<p style=" + quote + "text-align: justify;" + quote + ">--------------------------------------------------------------------------------------------------------------------------</p>" +
                                            "<p style=" + quote + "text-align: justify;" + quote + "><strong>" + txtNameProduct.Text.ToString() + "</strong></p>" +
                                            "<p style = " + quote + "text-align: justify;" + quote + "> " + txtContent.Text.ToString() + "</p>";
                rowHaravan.Cells[3].Value = ""; rowHaravan.Cells[4].Value = cbbNCC.Text.ToString(); rowHaravan.Cells[5].Value = txtLoaiSp.Text.ToString(); rowHaravan.Cells[6].Value = txtTag.Text.ToString(); rowHaravan.Cells[7].Value = "Yes"; rowHaravan.Cells[8].Value = "Title"; rowHaravan.Cells[9].Value = "Default Title"; rowHaravan.Cells[10].Value = ""; rowHaravan.Cells[11].Value = ""; rowHaravan.Cells[12].Value = ""; rowHaravan.Cells[13].Value = ""; rowHaravan.Cells[14].Value = txtSKU.Text.ToString(); rowHaravan.Cells[15].Value = "100"; rowHaravan.Cells[16].Value = "haravan"; rowHaravan.Cells[17].Value = txtSlHrv.Text.ToString(); rowHaravan.Cells[18].Value = "deny"; rowHaravan.Cells[19].Value = ""; rowHaravan.Cells[20].Value = txtPrice.Text.ToString(); rowHaravan.Cells[21].Value = txtPrice.Text.ToString(); rowHaravan.Cells[22].Value = "Yes"; rowHaravan.Cells[23].Value = "Yes"; rowHaravan.Cells[24].Value = txtSKU.Text.ToString(); rowHaravan.Cells[27].Value = "No"; rowHaravan.Cells[28].Value = txtNameProduct.Text.ToString();
                rowHaravan.Cells[29].Value = txtContent.Text.Substring(0, 70);
                rowHaravan.Cells[30].Value = ""; rowHaravan.Cells[31].Value = ""; rowHaravan.Cells[32].Value = ""; rowHaravan.Cells[33].Value = ""; rowHaravan.Cells[34].Value = ""; rowHaravan.Cells[35].Value = ""; rowHaravan.Cells[36].Value = ""; rowHaravan.Cells[37].Value = ""; rowHaravan.Cells[39].Value = ""; rowHaravan.Cells[40].Value = ""; rowHaravan.Cells[41].Value = DateTime.Now; rowHaravan.Cells[42].Value = DateTime.Now; rowHaravan.Cells[43].Value = "No"; rowHaravan.Cells[44].Value = "No"; rowHaravan.Cells[45].Value = "Yes"; rowHaravan.Cells[46].Value = "No"; rowHaravan.Cells[47].Value = "No"; rowHaravan.Cells[48].Value = "Yes";
                rowHaravan.Cells[38].Value = listCollection[0];
                rowHaravan.Cells[39].Value = listCollection[0];
                rowHaravan.Cells[25].Value = txtImg.Text.ToString();
                rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                dgvListHrv.Rows.Add(rowHaravan);
                txtSKU.Text = "";
                txtContent.Text = "";
                txtSize.Text = "";
                txtTrans.Text = "";
                txtNumpage.Text = "";
                txtPrice.Text = "";
                txtImg.Text = "";
                txtImg1.Text = "";
                txtImg2.Text = "";
                txtImg3.Text = "";
                txtImg4.Text = "";
                txtImg5.Text = "";
                txtImg6.Text = "";
                txtImg7.Text = "";
                txtImg8.Text = "";
                if (txtImg1.Text.ToString() != "" || listCollection[1] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg1.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg2.Text.ToString() != "" || listCollection[2] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg2.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg3.Text.ToString() != "" || listCollection[3] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg3.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg4.Text.ToString() != "" || listCollection[4] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg4.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg5.Text.ToString() != "" || listCollection[5] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg5.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg6.Text.ToString() != "" || listCollection[6] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg6.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if (txtImg7.Text.ToString() != "" || listCollection[7] != "")
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[25].Value = txtImg7.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
                if ((txtImg8.Text.ToString() != "") || (listCollection[8] != ""))
                {
                    index += 1;
                    rowHaravan = (DataGridViewRow)dgvListHrv.Rows[index].Clone();
                    rowHaravan.Cells[0].Value = txtNameProduct.Text.ToString();
                    rowHaravan.Cells[38].Value = listCollection[index];
                    rowHaravan.Cells[39].Value = listCollection[index];
                    rowHaravan.Cells[25].Value = txtImg8.Text.ToString();
                    rowHaravan.Cells[26].Value = txtNameProduct.Text.ToString();
                    dgvListHrv.Rows.Add(rowHaravan);
                }
            }
        }
        void addShopee(bool shopee)
        {
            if (shopee == false)
            {
                DataGridViewRow rowShopee = (DataGridViewRow)dgvListShopee.Rows[0].Clone();
                var breakLine = "\n";
                //string[] inDustryCode = cbbIndustry.Text.ToString().Split(" - ");
                List<string> inDustryCode = new List<string>(cbbIndustry.Text.ToString().Split(new string[] { " - " }, StringSplitOptions.None));
                rowShopee.Cells[0].Value = inDustryCode[0]; rowShopee.Cells[1].Value = "Sách - " + txtNameProduct.Text.ToString();
                //rowShopee.Cells[2].Value = txtContent.Text.ToString();
                rowShopee.Cells[2].Value = "Tên Nhà Cung Cấp: " + cbbNCC.Text.ToString() + breakLine +
                                            "Tác giả: " + cbbBrand.Text.ToString() + breakLine +
                                            "Người Dịch: " + txtTrans.Text.ToString() + breakLine +
                                            "NXB: " + txtNPH.Text.ToString() + breakLine +
                                            "Năm XB: " + txtNamsx.Text.ToString() + breakLine +
                                            "Ngôn Ngữ " + txtLanguage.Text.ToString() + breakLine +
                                            "Trọng lượng (gr): " + txtWeight.Text.ToString() + breakLine +
                                            "Kích Thước Bao Bì: " + txtSize.Text.ToString() + breakLine +
                                            "Số trang: " + txtNumpage.Text.ToString() + breakLine +
                                            "Hình thức: " + cbbLoaiBia.Text.ToString() + breakLine +
                                            ">--------------------------------------------------------------------------------------------------------------------------<" + breakLine +
                                            txtNameProduct.Text.ToString() + breakLine +
                                            txtContent.Text.ToString();
                rowShopee.Cells[3].Value = txtSKU.Text.ToString(); rowShopee.Cells[4].Value = ""; rowShopee.Cells[5].Value = ""; rowShopee.Cells[6].Value = ""; rowShopee.Cells[7].Value = ""; rowShopee.Cells[8].Value = ""; rowShopee.Cells[9].Value = ""; rowShopee.Cells[10].Value = txtPrice.Text.ToString(); rowShopee.Cells[11].Value = txtSlShopee.Text.ToString(); rowShopee.Cells[12].Value = ""; rowShopee.Cells[13].Value = ""; rowShopee.Cells[14].Value = txtImg.Text.ToString(); rowShopee.Cells[15].Value = txtImg1.Text.ToString(); rowShopee.Cells[16].Value = txtImg2.Text.ToString(); rowShopee.Cells[17].Value = txtImg3.Text.ToString(); rowShopee.Cells[18].Value = txtImg4.Text.ToString(); rowShopee.Cells[19].Value = txtImg5.Text.ToString(); rowShopee.Cells[20].Value = txtImg6.Text.ToString(); rowShopee.Cells[21].Value = txtImg7.Text.ToString(); rowShopee.Cells[22].Value = txtImg8.Text.ToString(); rowShopee.Cells[23].Value = txtWeight.Text.ToString(); rowShopee.Cells[24].Value = ""; rowShopee.Cells[25].Value = ""; rowShopee.Cells[26].Value = ""; rowShopee.Cells[27].Value = "Bật"; rowShopee.Cells[28].Value = ""; rowShopee.Cells[29].Value = "=IFERROR(VLOOKUP(@INDIRECT(ADDRESS(ROW(),1)),mp_advanced_filter_prohibit_cat!$A$1:$EA$20000,2,0)," + " " + ")"; rowShopee.Cells[30].Value = txtIdBrand.Text.ToString(); rowShopee.Cells[31].Value = txtXuatKhau.Text.ToString(); rowShopee.Cells[32].Value = ""; rowShopee.Cells[33].Value = ""; rowShopee.Cells[34].Value = txtNPH.Text.ToString(); rowShopee.Cells[35].Value = ""; rowShopee.Cells[36].Value = txtLanguage.Text.ToString(); rowShopee.Cells[37].Value = txtXuatKhau.Text.ToString(); rowShopee.Cells[38].Value = ""; rowShopee.Cells[39].Value = ""; rowShopee.Cells[40].Value = ""; rowShopee.Cells[41].Value = ""; rowShopee.Cells[42].Value = cbbLoaiBia.Text.ToString(); rowShopee.Cells[43].Value = txtNamsx.Text.ToString(); rowShopee.Cells[44].Value = ""; rowShopee.Cells[45].Value = ""; rowShopee.Cells[46].Value = "";
                dgvListShopee.Rows.Add(rowShopee);
            }
        }
        void addTiktok(bool tiktok)
        {
            if (tiktok == false)
            {
                DataGridViewRow rowTikTok = (DataGridViewRow)dgvTikTok.Rows[0].Clone();
                var breakLine = "\n";
                //string[] inDustryCode = cbbIndustry.Text.ToString().Split(" - ");
                List<string> inDustryCode = new List<string>(cbbIndustry.Text.ToString().Split(new string[] { " - " }, StringSplitOptions.None));
                rowTikTok.Cells[0].Value = inDustryCode[0]; rowTikTok.Cells[1].Value = "7258118283534665477/FABICO";
                rowTikTok.Cells[2].Value = txtNameProduct.Text.ToString();
                rowTikTok.Cells[3].Value = "Tên Nhà Cung Cấp: " + cbbNCC.Text.ToString() + breakLine +
                                            "Tác giả: " + cbbBrand.Text.ToString() + breakLine +
                                            "Người Dịch: " + txtTrans.Text.ToString() + breakLine +
                                            "NXB: " + txtNPH.Text.ToString() + breakLine +
                                            "Năm XB: " + txtNamsx.Text.ToString() + breakLine +
                                            "Ngôn Ngữ " + txtLanguage.Text.ToString() + breakLine +
                                            "Trọng lượng (gr): " + txtWeight.Text.ToString() + breakLine +
                                            "Kích Thước Bao Bì: " + txtSize.Text.ToString() + breakLine +
                                            "Số trang: " + txtNumpage.Text.ToString() + breakLine +
                                            "Hình thức: " + cbbLoaiBia.Text.ToString() + breakLine +
                                            ">--------------------------------------------------------------------------------------------------------------------------<" + breakLine +
                                            txtNameProduct.Text.ToString() + breakLine +
                                            txtContent.Text.ToString();
                rowTikTok.Cells[3].Value = txtWeight.Text.ToString(); rowTikTok.Cells[4].Value = ""; rowTikTok.Cells[5].Value = ""; rowTikTok.Cells[6].Value = ""; rowTikTok.Cells[7].Value = ""; rowTikTok.Cells[8].Value = ""; rowTikTok.Cells[9].Value = ""; rowTikTok.Cells[10].Value = ""; rowTikTok.Cells[11].Value = ""; rowTikTok.Cells[12].Value = txtPrice.Text.ToString(); rowTikTok.Cells[13].Value = txtSlTikTok.Text.ToString(); rowTikTok.Cells[14].Value = ""; rowTikTok.Cells[15].Value = txtImg.Text.ToString(); rowTikTok.Cells[16].Value = txtImg1.Text.ToString(); rowTikTok.Cells[17].Value = txtImg2.Text.ToString(); rowTikTok.Cells[18].Value = txtImg3.Text.ToString(); rowTikTok.Cells[19].Value = txtImg4.Text.ToString(); rowTikTok.Cells[20].Value = txtImg5.Text.ToString(); rowTikTok.Cells[21].Value = txtImg6.Text.ToString(); rowTikTok.Cells[22].Value = txtImg7.Text.ToString(); rowTikTok.Cells[23].Value = txtImg8.Text.ToString(); rowTikTok.Cells[24].Value = ""; rowTikTok.Cells[25].Value = txtNamsx.Text.ToString(); rowTikTok.Cells[26].Value = txtNPH.Text.ToString(); rowTikTok.Cells[27].Value =txtISBN.Text.ToString(); rowTikTok.Cells[28].Value = txtTrans.Text.ToString(); rowTikTok.Cells[29].Value = ""; rowTikTok.Cells[30].Value = txtNumpage.Text.ToString();
                dgvTikTok.Rows.Add(rowTikTok);
            }
        }
        private void btnAddProduct_Click(object sender, EventArgs e)
        {
            try
            {
                //dtgvInfo.Rows.Add(txtNameProduct.Text.ToString(), txtContent.Text.ToString(), txtIdBrand.Text.ToString(), txtXuatKhau.Text.ToString(), txtLang.Text.ToString(), txtNXB.Text.ToString(), txtLoaiBia.Text.ToString(), txtNamsx.Text.ToString(), txtPrice.Text.ToString(), txtslTon.Text.ToString(), txtWeight.Text.ToString(), txtSKU.Text.ToString(), txtImg.Text.ToString());
                
                //--------------------------Add Datagridview cho Shopee-----------------------------------------------------------------------//
                addTiktok(false);


                //--------------------------Add Datagridview cho Shopee-----------------------------------------------------------------------//
                addShopee(false);

                //--------------------------Add Datagridview cho Haravan-----------------------------------------------------------------------//
                addHaravan(false);

              

                clear();
                //-------------------------------------------------------------------------------------------------//
            }
            catch {  }
        }

        void clear()
        {
            txtNameProduct.Text = "";
            txtContent.Text = "";
            txtIdBrand.Text = "";
            //txtNXB.Text = "";
            txtPrice.Text = "";
            txtSlShopee.Text = "";
            txtSlHrv.Text = "";
            txtSKU.Text = "";
            txtImg.Text = ""; txtImg1.Text = ""; txtImg2.Text = ""; txtImg3.Text = ""; txtImg4.Text = ""; txtImg5.Text = ""; txtImg6.Text = ""; txtImg7.Text = ""; txtImg8.Text = "";
        }
        public class Incident
        {
            public Incident(){}
            public Incident(string id, string names, string display_names)
            {
                brand_id = id;
                name = names;
                display_name = display_names;
            }
            public string brand_id { get; set; }
            public string name { get; set; }
            public string display_name { get; set; }
        }

        public class InDustry
        {
            public InDustry() { }
            public InDustry(string idIndustry, string namesIndustry)
            {
                id_Industry = idIndustry;
                Industry_name = namesIndustry;
            }
            public string id_Industry { get; set; }
            public string Industry_name { get; set; }
        }
        void exportExcelHaravan()
        {
            // creating Excel Application  
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            _Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            _Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Sheet1";
            // storing header part in Excel  
            for (int i = 1; i < dgvListHrv.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dgvListHrv.Columns[i - 1].HeaderText;

            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgvListHrv.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dgvListHrv.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dgvListHrv.Rows[i].Cells[j].Value?.ToString();
                }
            }
            //app.DisplayAlerts = false;
            app.AskToUpdateLinks = false;
            app.DisplayAlerts = false;
            // save the application  
            workbook.SaveAs("d:\\xuatharavanvashopee\\haravan.xls", XlFileFormat.xlExcel8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Exit from the application  
            app.Quit();
        }
        void exportExcelShopee()
        {
            // creating Excel Application  
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            _Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            _Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            app.AlertBeforeOverwriting = false;
            app.DisplayAlerts = false;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            // get the reference of first sheet. By default its name is Sheet2.  

            worksheet.Name = @"Phạm vi ngày chuẩn bị hàng";
            /* worksheet.Cells[1, 1] = "100643 - ";*/
            worksheet.Cells[1, 1] = "et_title_category_name";
            worksheet.Cells[1, 2] = "et_title_category_id";
            worksheet.Cells[1, 3] = "et_title_dts_range";
            worksheet.Cells[1, 4] = "et_title_remark";

            // get the reference of first sheet. By default its name is Sheet3.
            Worksheet oSheet3 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            oSheet3.Name = @"Đăng tải bản mẫu";
            oSheet3.Cells[1, 1] = "ps_category"; oSheet3.Cells[1, 2] = "ps_product_name"; oSheet3.Cells[1, 3] = "ps_product_description"; oSheet3.Cells[1, 4] = "ps_sku_parent_short"; oSheet3.Cells[1, 5] = "et_title_variation_integration_no"; oSheet3.Cells[1, 6] = "et_title_variation_1"; oSheet3.Cells[1, 7] = "et_title_option_for_variation_1"; oSheet3.Cells[1, 8] = "et_title_image_per_variation"; oSheet3.Cells[1, 9] = "et_title_variation_2"; oSheet3.Cells[1, 10] = "et_title_option_for_variation_2"; oSheet3.Cells[1, 11] = "ps_price"; oSheet3.Cells[1, 12] = "ps_stock"; oSheet3.Cells[1, 13] = "ps_sku_short"; oSheet3.Cells[1, 14] = "ps_hs_code"; oSheet3.Cells[1, 15] = "ps_tax_code"; oSheet3.Cells[1, 16] = "ps_item_cover_image"; oSheet3.Cells[1, 17] = "ps_item_image_1"; oSheet3.Cells[1, 18] = "ps_item_image_2"; oSheet3.Cells[1, 19] = "ps_item_image_3"; oSheet3.Cells[1, 20] = "ps_item_image_4"; oSheet3.Cells[1, 21] = "ps_item_image_5"; oSheet3.Cells[1, 22] = "ps_item_image_6"; oSheet3.Cells[1, 23] = "ps_item_image_7"; oSheet3.Cells[1, 24] = "ps_item_image_8"; oSheet3.Cells[1, 25] = "ps_weight"; oSheet3.Cells[1, 26] = "ps_length"; oSheet3.Cells[1, 27] = "ps_width"; oSheet3.Cells[1, 28] = "ps_height"; oSheet3.Cells[1, 29] = "channel_id_"; oSheet3.Cells[1, 30] = "channel_id_"; oSheet3.Cells[1, 31] = "channel_id_"; oSheet3.Cells[1, 32] = "ps_product_pre_order_dts_range"; oSheet3.Cells[1, 33] = "ps_product_pre_order_dts"; oSheet3.Cells[1, 34] = "ps_brand"; oSheet3.Cells[1, 35] = ""; oSheet3.Cells[1, 36] = ""; oSheet3.Cells[1, 37] = ""; oSheet3.Cells[1, 38] = ""; oSheet3.Cells[1, 39] = ""; oSheet3.Cells[1, 40] = ""; oSheet3.Cells[1, 41] = ""; oSheet3.Cells[1, 42] = ""; oSheet3.Cells[1, 43] = ""; oSheet3.Cells[1, 44] = "ps_tool_mass_upload_sample_attr_country_origin"; oSheet3.Cells[1, 45] = "ps_tool_mass_upload_sample_attr_manufacturer_details"; oSheet3.Cells[1, 46] = "ps_tool_mass_upload_sample_attr_packer_details"; oSheet3.Cells[1, 47] = "ps_tool_mass_upload_sample_attr_importer_details";
            oSheet3.Cells[2, 1] = "basic";
            // Sheet2
            Worksheet oSheet2 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            oSheet2.Name = @"Bản đăng tải";
            oSheet2.Cells[1, 1] = "ps_category"; oSheet2.Cells[1, 2] = "ps_product_name"; oSheet2.Cells[1, 3] = "ps_product_description"; oSheet2.Cells[1, 4] = "ps_sku_parent_short"; oSheet2.Cells[1, 5] = "et_title_variation_integration_no"; oSheet2.Cells[1, 6] = "et_title_variation_1"; oSheet2.Cells[1, 7] = "et_title_option_for_variation_1"; oSheet2.Cells[1, 8] = "et_title_image_per_variation"; oSheet2.Cells[1, 9] = "et_title_variation_2"; oSheet2.Cells[1, 10] = "et_title_option_for_variation_2"; oSheet2.Cells[1, 11] = "ps_price"; oSheet2.Cells[1, 12] = "ps_stock"; oSheet2.Cells[1, 13] = "ps_sku_short"; oSheet2.Cells[1, 14] = "ps_new_size_chart"; oSheet2.Cells[1, 15] = "ps_item_cover_image"; oSheet2.Cells[1, 16] = "ps_item_image_1"; oSheet2.Cells[1, 17] = "ps_item_image_2"; oSheet2.Cells[1, 18] = "ps_item_image_3"; oSheet2.Cells[1, 19] = "ps_item_image_4"; oSheet2.Cells[1, 20] = "ps_item_image_5"; oSheet2.Cells[1, 21] = "ps_item_image_6"; oSheet2.Cells[1, 22] = "ps_item_image_7"; oSheet2.Cells[1, 23] = "ps_item_image_8"; oSheet2.Cells[1, 24] = "ps_weight"; oSheet2.Cells[1, 25] = "ps_length"; oSheet2.Cells[1, 26] = "ps_width"; oSheet2.Cells[1, 27] = "ps_height"; oSheet2.Cells[1, 28] = "channel_id.5001"; oSheet2.Cells[1, 29] = "ps_product_pre_order_dts"; oSheet2.Cells[1, 30] = "ps_product_pre_order_dts_range"; oSheet2.Cells[1, 31] = "ps_brand"; oSheet2.Cells[1, 32] = "ps_product_global_attribute.100037"; oSheet2.Cells[1, 33] = "ps_product_global_attribute.100121"; oSheet2.Cells[1, 34] = "ps_product_global_attribute.100370"; oSheet2.Cells[1, 35] = "ps_product_global_attribute.100669"; oSheet2.Cells[1, 36] = "ps_product_global_attribute.100670"; oSheet2.Cells[1, 37] = "ps_product_global_attribute.100673"; oSheet2.Cells[1, 38] = "ps_product_global_attribute.100676"; oSheet2.Cells[1, 39] = "ps_product_global_attribute.100691"; oSheet2.Cells[1, 40] = "ps_product_global_attribute.100697"; oSheet2.Cells[1, 41] = "ps_product_global_attribute.100707"; oSheet2.Cells[1, 42] = "ps_product_global_attribute.100709"; oSheet2.Cells[1, 43] = "ps_product_global_attribute.100710"; oSheet2.Cells[1, 44] = "ps_product_global_attribute.101024"; oSheet2.Cells[1, 45] = "ps_product_global_attribute.101059"; oSheet2.Cells[1, 46] = "ps_product_global_attribute.101067"; oSheet2.Cells[1, 47] = "ps_product_global_attribute.101068"; oSheet2.Cells[1, 48] = "et_title_reason";
            oSheet2.Cells[2, 1] = "advanced"; oSheet2.Cells[2, 2] = "9341b077d91d983cf5e0b3ec7c4dd24b"; oSheet2.Cells[2, 3] = "100643";
            // storing header part in Excel
            for (int i = 1; i < dgvListShopee.Columns.Count + 1; i++)
            {
                oSheet2.Cells[3, i] = dgvListShopee.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgvListShopee.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dgvListShopee.Columns.Count; j++)
                {
                    oSheet2.Cells[i + 7, j + 1] = dgvListShopee.Rows[i].Cells[j].Value?.ToString();
                }
            }

            // Sheet4
            Worksheet oSheet4 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            oSheet4.Name = @"Hướng dẫn";
            /*
            oSheet4.Cells[2, 1] = "advance";
            oSheet4.Cells[2, 2] = "220309_sizechart";
            */
            // save the application  

            app.AskToUpdateLinks = false;
            //app.DisplayAlerts = false;
            workbook.SaveAs("d:\\xuatharavanvashopee\\shopee.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            // Exit from the application  
            app.Quit();
        }

        void exportExcelTikTok()
        {
            // creating Excel Application  
            _Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            _Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            _Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            app.AlertBeforeOverwriting = false;
            app.DisplayAlerts = false;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            // get the reference of first sheet. By default its name is Sheet2.  

            worksheet.Name = @"Instruction";
          
            // get the reference of first sheet. By default its name is Sheet3.
            Worksheet oSheet3 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            oSheet3.Name = @"Example";
         
            // Sheet2
            Worksheet oSheet2 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            oSheet2.Name = @"Template";
            oSheet2.Cells[1, 1] = "category"; oSheet2.Cells[1, 2] = "brand"; oSheet2.Cells[1, 3] = "product_name"; oSheet2.Cells[1, 4] = "product_description"; oSheet2.Cells[1, 5] = "parcel_weight"; oSheet2.Cells[1, 6] = "parcel_length"; oSheet2.Cells[1, 7] = "parcel_width"; oSheet2.Cells[1, 8] = "parcel_height"; oSheet2.Cells[1, 9] = "delivery"; oSheet2.Cells[1, 10] = "property_value_1"; oSheet2.Cells[1, 11] = "property_1_image"; oSheet2.Cells[1, 12] = "property_value_2"; oSheet2.Cells[1, 13] = "price"; oSheet2.Cells[1, 14] = "quantity"; oSheet2.Cells[1, 15] = "seller_sku"; oSheet2.Cells[1, 16] = "main_image"; oSheet2.Cells[1, 17] = "image_2"; oSheet2.Cells[1, 18] = "image_3"; oSheet2.Cells[1, 19] = "image_4"; oSheet2.Cells[1, 20] = "image_5"; oSheet2.Cells[1, 21] = "image_6"; oSheet2.Cells[1, 22] = "image_7"; oSheet2.Cells[1, 23] = "image_8"; oSheet2.Cells[1, 24] = "image_9"; oSheet2.Cells[1, 25] = "size_chart"; oSheet2.Cells[1, 26] = "product_property/100530"; oSheet2.Cells[1, 27] = "product_property/100532"; oSheet2.Cells[1, 28] = "product_property/100534"; oSheet2.Cells[1, 29] = "product_property/100536"; oSheet2.Cells[1, 30] = "product_property/100537"; oSheet2.Cells[1, 31] = "product_property/100538";
            oSheet2.Cells[2, 1] = "V2"; oSheet2.Cells[2, 2] = "create_product"; oSheet2.Cells[2, 3] = "metric";
            // storing header part in Excel
            for (int i = 1; i < dgvTikTok.Columns.Count + 1; i++)
            {
                oSheet2.Cells[3, i] = dgvTikTok.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgvTikTok.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dgvTikTok.Columns.Count; j++)
                {
                    oSheet2.Cells[i + 7, j + 1] = dgvTikTok.Rows[i].Cells[j].Value?.ToString();
                }
            }



            // save the application  

            app.AskToUpdateLinks = false;
            //app.DisplayAlerts = false;
            workbook.SaveAs("d:\\xuatharavanvashopee\\tiktok.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            // Exit from the application  
            app.Quit();
        }


        private void ExportShopeeAndHaravan_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void btnGetId_Click(object sender, EventArgs e)
        {
            try
            {
                // Read values from file
                var strReadJson = File.ReadAllText(@"id-brand.json");
                // Convert to Json Object
                var bx = JsonConvert.DeserializeObject<List<Incident>>(strReadJson);
                if ((cbbBrand.SelectedIndex > -1) || (cbbBrand.Text != null))
                {
                    txtIdBrand.Text = bx.Where(x => x.name == cbbBrand.Text.ToString()).Select(x => x.brand_id.ToString()).FirstOrDefault();
                }
                else
                {
                    MessageBox.Show("Không có thông tin", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {

            }
        }

        private void txtNameProduct_TextChanged(object sender, EventArgs e)
        {
            int count = txtNameProduct.TextLength;
            lbCount.Text = count.ToString();
            if (count > 113)
            {
                lbCount.ForeColor = Color.Red;
            }
            else
            {
                lbCount.ForeColor = Color.Black;
            }

            if (txtNameProduct.Text != null)
            {
                List<string> tags = new List<string>(txtNameProduct.Text.ToString().Split(new string[] { " - " }, StringSplitOptions.None));
                txtTag.Text = tags[0];
            }
        }

        private void btnClearRow_Click(object sender, EventArgs e)
        {
            dgvListShopee.Rows.Clear();
            dgvListHrv.Rows.Clear();
        }

        private void btnExportHaravan_Click(object sender, EventArgs e)
        {
            try
            {
                exportExcelHaravan();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Thông Báo",MessageBoxButtons.OK,MessageBoxIcon.Error);
                throw;
            }
        }

        private void btnExportShopee_Click(object sender, EventArgs e)
        {
            try
            {
                exportExcelShopee();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void btnExportAll_Click(object sender, EventArgs e)
        {
            try
            {
                exportExcelShopee();
                exportExcelHaravan();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
          
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Main mains = new Main();
            mains.Show();
            this.Hide();
        }

        private void btnExportTiktok_Click(object sender, EventArgs e)
        {
            try
            {
                exportExcelTikTok();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }
}
