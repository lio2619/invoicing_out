using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Invoicing.form
{
    public partial class 所有商品 : Form
    {
        bool s = false; //false = 貨品編號  true=商品名稱
        public 所有商品()
        {
            InitializeComponent();
            string SQL = "select 貨品編號, 品名, 標準售價, 售價A, 售價B, 標準成本 from 貨品主檔 order by 貨品編號";
            SQLConnect con = new SQLConnect();
            dataGridView1.DataSource = con.Find(SQL);
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "商品名稱")
                s = true;
            else
                s = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DataTable DT = (DataTable)dataGridView1.DataSource;
            string search = textBox1.Text.Trim().ToString();
            if (s)
            {
                var row = dataGridView1.Rows.Cast<DataGridViewRow>()
                        .Where(x => !x.IsNewRow)
                        .Where(x => ((DataRowView)x.DataBoundItem)["品名"].ToString().StartsWith(search))
                        .FirstOrDefault();
                if (row == null)
                {
                    MessageBox.Show("沒有這個品名");
                    return;
                }
                dataGridView1.CurrentCell = row.Cells[0];
            }
            else
            {
                var row = dataGridView1.Rows.Cast<DataGridViewRow>()
                        .Where(x => !x.IsNewRow)
                        .Where(x => ((DataRowView)x.DataBoundItem)["貨品編號"].ToString().StartsWith(search))
                        .FirstOrDefault();
                if (row == null)
                {
                    MessageBox.Show("沒有這個編號");
                    return;
                }
                dataGridView1.CurrentCell = row.Cells[0];
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string find = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            Form[] aryf = this.Parent.FindForm().MdiChildren;
            foreach (Form f in aryf)
            {
                if (f.Name == "管理產品")
                {
                    ((管理產品)f).search(find);
                    break;
                }
            }
        }
    }
}
