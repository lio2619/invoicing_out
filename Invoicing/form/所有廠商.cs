using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Invoicing.form
{
    public partial class 所有廠商 : Form
    {
        public 所有廠商()
        {
            InitializeComponent();
            string SQL = "select 公司編號 ,公司全名, 聯絡電話一, 傳真號碼, 送貨地址 from 廠商 order by 公司全名";
            SQLConnect con = new SQLConnect();
            dataGridView1.DataSource = con.Find(SQL);
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoGenerateColumns = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string search = textBox1.Text.Trim().ToString();
            var row = dataGridView1.Rows.Cast<DataGridViewRow>()
                        .Where(x => !x.IsNewRow)
                        .Where(x => ((DataRowView)x.DataBoundItem)["公司全名"].ToString().Contains(search))
                        .FirstOrDefault();
            if (row == null)
            {
                MessageBox.Show("沒有這個公司");
                return;
            }
            dataGridView1.CurrentCell = row.Cells[0];
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string find = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            string number = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            Form[] aryf = this.Parent.FindForm().MdiChildren;
            foreach (Form f in aryf)
            {
                if (f.Name == "管理廠商")
                {
                    ((管理廠商)f).search(find, number);
                    break;
                }
            }
        }
    }
}
