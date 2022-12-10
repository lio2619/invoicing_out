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
            string SQL = "select * from 廠商 order by 公司全名";
            SQLConnect con = new SQLConnect();
            dataGridView1.DataSource = con.Find(SQL);
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
    }
}
