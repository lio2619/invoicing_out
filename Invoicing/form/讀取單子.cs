using System;
using System.Data;
using System.Globalization;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 讀取單子 : Form
    {
        private string NAME = "";
        private string number = "";
        private string STORE = "";
        private string NOTE = "";
        private string TOTAL = "";
        private string DATE = "";
        private string BILL = "";
        public 讀取單子()
        {
            InitializeComponent();
        }
        public 讀取單子(string name)
        {
            InitializeComponent();
            NAME = name;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select 單子編號 as 編號, 日期, 客戶, 單子, 總金額, 備註 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子=@bili and 刪除='0' order by 客戶";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                bili = NAME
            });
            dataGridView1.Columns.Clear();
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView1.Columns[e.ColumnIndex].Name == "select" && e.RowIndex >= 0)
                {
                    number = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();      //單子編號
                    DATE = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();        //日期
                    STORE = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();       //客戶
                    BILL = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();        //單子
                    NOTE = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();        //備註
                    TOTAL = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();       //總金額
                }
            }
            catch
            {
                MessageBox.Show("請把讀檔的視窗便在開啟來一次");
                return;
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            DataTable DT;
            SQLConnect con = new SQLConnect();
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            btn.Name = "search";
            btn.HeaderText = "查詢";
            btn.DefaultCellStyle.NullValue = "查詢";
            string SQL = "";
            if (number == "") return;
            else
            {
                SQL = "select 貨品編號, 品名, 數量, 基本單位, 單價, 金額, 備註" +
                    " from 整張儲存 where 單子編號=@number";
                DT = await con.searchDataTable(SQL, new { number = number});
            }
            if (NAME == "採購單")
            {
                SQL = "select 貨品編號, 品名, 數量, 基本單位, 備註" +
                    " from 整張儲存 where 單子編號=@number";
                DT = await con.searchDataTable(SQL, new { number = number });
                採購單 form2 = (採購單)Owner;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            else if (NAME == "出貨退出單")
            {
                出貨退出單 form2 = (出貨退出單)Owner;
                form2.temp = DT.Rows.Count;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.label2.Text = TOTAL;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            else if (NAME == "出貨單")
            {
                出貨單 form2 = (出貨單)Owner;
                form2.temp = DT.Rows.Count;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.label2.Text = TOTAL;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            else if (NAME == "訂貨單")
            {
                訂貨單 form2 = (訂貨單)Owner;
                form2.temp = DT.Rows.Count;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.label2.Text = TOTAL;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            else if (NAME == "進貨退出單")
            {
                進貨退出單 form2 = (進貨退出單)Owner;
                form2.temp = DT.Rows.Count;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.label2.Text = TOTAL;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            else if (NAME == "進貨單")
            {
                進貨單 form2 = (進貨單)Owner;
                form2.temp = DT.Rows.Count;
                form2.dataGridView1.AutoGenerateColumns = true;
                form2.label6.Text = number;
                form2.dateTimePicker1.Value = DateTime.ParseExact(DATE, "yyyyMMdd", CultureInfo.InvariantCulture);
                form2.comboBox1.Text = STORE;
                form2.textBox1.Text = NOTE;
                form2.label2.Text = TOTAL;
                form2.dataGridView1.Columns.Clear();
                form2.dataGridView1.DataSource = DT;
                form2.dataGridView1.AutoResizeColumns();
                form2.dataGridView1.Columns.Add(btn);
                form2.dataGridView1.AutoGenerateColumns = false;
                form2.label6.Visible = true;
            }
            Close();
        }
    }
}
