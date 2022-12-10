using System.Data;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 總貨品 : Form
    {
        public 總貨品()
        {
            InitializeComponent();
            SQLConnect con = new SQLConnect();
            string SQL = "select 貨品編號, 品名, 標準售價, 售價A, 售價B, 標準成本 from 貨品主檔 order by 貨品編號";
            DataTable DT = con.Find(SQL);
            dataGridView1.DataSource = DT;
            dataGridView1.AutoResizeColumns();
        }

        private void textBox1_TextChanged(object sender, System.EventArgs e)
        {
        }
    }
}
