using System;
using System.Data;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 總金額 : Form
    {
        public 總金額()
        {
            InitializeComponent();
        }
        //出貨、訂貨、出貨退出
        private async void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            label4.Text = "0";
            SQLConnect con = new SQLConnect();
            string SQL = "select 客戶, 單子, sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 in('出貨退出單', '出貨單')" +
                        " group by 客戶, 單子 order by 客戶";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd")
            });
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                }
            }
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 ='出貨退出單'";
            dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd")
            });
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 ='出貨單'";
            DataTable DT = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd")
            });
            if (dt != null && DT != null)
            {
                if (dt.Rows.Count > 0 && DT.Rows.Count > 0)
                {
                    if (DT.Rows[0]["金額"].ToString() != "" && dt.Rows[0]["金額"].ToString() != "")
                        label4.Text = (double.Parse(DT.Rows[0]["金額"].ToString()) - double.Parse(dt.Rows[0]["金額"].ToString())).ToString();
                    else if (DT.Rows[0]["金額"].ToString() != "" && dt.Rows[0]["金額"].ToString() == "")
                        label4.Text = DT.Rows[0]["金額"].ToString();
                    else if (DT.Rows[0]["金額"].ToString() == "" && dt.Rows[0]["金額"].ToString() != "")
                        label4.Text = dt.Rows[0]["金額"].ToString();
                }
            }
        }
    }
}
