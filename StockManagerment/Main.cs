using System;
using System.Windows.Forms;
namespace StockManagerment
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        private void btnStockManagerment_Click(object sender, EventArgs e)
        {
            StockManagerment stockManagerment = new StockManagerment();
            stockManagerment.Show();
            this.Hide();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            ExportShopeeAndHaravan exportShopeeAndHaravan = new ExportShopeeAndHaravan();
            exportShopeeAndHaravan.Show();
            this.Hide();
        }

        private void btninsertupdateShopee_Click(object sender, EventArgs e)
        {
            SearchForm searchForm = new SearchForm();
            searchForm.Show();
            this.Hide();
        }

        private void btnTikTok_Click(object sender, EventArgs e)
        {
            TikTokInfo tiktokform = new TikTokInfo();
            tiktokform.Show();
            this.Hide();
        }
    }
}
