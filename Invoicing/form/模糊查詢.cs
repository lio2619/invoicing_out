using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 模糊查詢 : Form
    {
        private int POINT = 0;
        private string number = "";
        private string NAME = "";
        public 模糊查詢()
        {
            InitializeComponent();
        }
        public 模糊查詢(int point, string name)
        {
            InitializeComponent();
            this.POINT = point;
            this.NAME = name;
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                                                   dataGridView1.RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dataGridView1.RowHeadersDefaultCellStyle.Font,
                                   rectangle, dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            SQLConnect con = new SQLConnect();
            string sql = "select 貨品編號, 品名, 基本單位, 標準售價 from 貨品主檔 where 貨品編號 like @name & '%' order by 貨品編號";
            DataTable dt = await con.searchDataTable(sql, new {name = textBox1.Text.Trim()});
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                    checkBoxColumn.Name = "select";
                    checkBoxColumn.HeaderText = "選擇";
                    dataGridView1.Columns.Add(checkBoxColumn);
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (number == "") return;
            if (NAME == "採購單")
            {
                採購單 form2 = (採購單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            else if (NAME == "出貨退出單")
            {
                出貨退出單 form2 = (出貨退出單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            else if (NAME == "出貨單")
            {
                出貨單 form2 = (出貨單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            else if (NAME == "訂貨單")
            {
                訂貨單 form2 = (訂貨單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            else if (NAME == "進貨退出單")
            {
                進貨退出單 form2 = (進貨退出單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            else if (NAME == "進貨單")
            {
                進貨單 form2 = (進貨單)Owner;
                form2.dataGridView1.Rows[POINT].Cells[0].Value = number;
                form2.dataGridView1.CurrentCell = form2.dataGridView1.Rows[POINT].Cells[0];
                form2.dataGridView1.BeginEdit(true);
            }
            Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "select" && e.RowIndex >= 0)
            {
                number = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            }
        }
    }
}
