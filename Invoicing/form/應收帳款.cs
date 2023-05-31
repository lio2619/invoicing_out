using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Invoicing.form
{
    public partial class 應收帳款 : Form
    {
        public 應收帳款()
        {
            InitializeComponent();
        }

        private async void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            列印排版 form = new 列印排版();
            SQLConnect con = new SQLConnect();
            string SQL = "select 客戶, 總金額, 單子, 日期, 單子編號 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 in('出貨退出單', '出貨單')" +
                        "and 客戶=@customer and 刪除='0' order by 單子";
            DataTable DT = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            form.collect_money_header(e, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            int high = 165;
            int row = 0;
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                row += 1;
                int no = int.Parse(DT.Rows[i]["單子編號"].ToString());
                string uid = DT.Rows[i]["日期"].ToString() + no.ToString("000");
                form.collect_money_inside(e, DT.Rows[i]["單子"].ToString(), high, 0);
                form.collect_money_inside(e, DT.Rows[i]["日期"].ToString(), high, 1);
                form.collect_money_inside(e, uid, high, 2);
                if(DT.Rows[i]["單子"].ToString() == "出貨退出單")
                {
                    form.collect_money_inside(e, "- " + DT.Rows[i]["總金額"].ToString(), high, 3);
                    form.collect_money_inside(e, "- " + DT.Rows[i]["總金額"].ToString(), high, 4);
                }
                else
                {
                    form.collect_money_inside(e, DT.Rows[i]["總金額"].ToString(), high, 3);
                    form.collect_money_inside(e, DT.Rows[i]["總金額"].ToString(), high, 4);
                }
                
                high += 20;
            }
            Pen black = new Pen(Color.Black);
            PointF point1 = new PointF(20, high);
            PointF point2 = new PointF(790, high);
            e.Graphics.DrawLine(black, point1, point2);
            high += 10;

            string total = "";
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 ='出貨退出單' " +
                    "and 客戶=@customer and 刪除='0'";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day " +
                    "and 單子 ='出貨單' and 客戶=@customer and 刪除='0'";
            DT = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            if (dt != null && DT != null)
            {
                if (dt.Rows.Count > 0 && DT.Rows.Count > 0)
                {
                    if (DT.Rows[0]["金額"].ToString() != "" && dt.Rows[0]["金額"].ToString() != "")
                        total = (double.Parse(DT.Rows[0]["金額"].ToString()) - double.Parse(dt.Rows[0]["金額"].ToString())).ToString();
                    else if (DT.Rows[0]["金額"].ToString() != "" && dt.Rows[0]["金額"].ToString() == "")
                        total = DT.Rows[0]["金額"].ToString();
                    else if (DT.Rows[0]["金額"].ToString() == "" && dt.Rows[0]["金額"].ToString() != "")
                        total = "-" + dt.Rows[0]["金額"].ToString();
                }
            }
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("本期合計：", new Font("Arial", 12), Brushes.Black, new Point(530, high));
            e.Graphics.DrawString(total, new Font("Arial", 12), Brushes.Black, new Point(720, high), stringFormat);
            high += 20;
            e.Graphics.DrawString("營業稅：", new Font("Arial", 12), Brushes.Black, new Point(530, high));
            e.Graphics.DrawString(textBox1.Text.Trim().ToString(), new Font("Arial", 12), Brushes.Black, new Point(720, high), stringFormat);
            high += 20;
            decimal cost = decimal.Parse(total) + decimal.Parse(textBox1.Text.Trim().ToString());
            e.Graphics.DrawString("本期總計：", new Font("Arial", 12), Brushes.Black, new Point(530, high));
            e.Graphics.DrawString(cost.ToString(), new Font("Arial", 12), Brushes.Black, new Point(720, high), stringFormat);

            //form.collect_money_last(e, row);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("請選擇客戶");
                return;
            }
            if (textBox1.Text.ToString() == "")
            {
                MessageBox.Show("請輸入營業稅");
                return;
            }
            decimal number;
            bool ret = decimal.TryParse(textBox1.Text.ToString(), out number);
            if (!ret)
            {
                MessageBox.Show("營業稅請輸入數值");
                return;
            }
            printPreviewDialog1.Document = printDocument1;
            try
            {
                short s = short.Parse(Interaction.InputBox("請輸入要印幾份", "標題", "1", -1, -1));
                printDocument1.PrinterSettings.Copies = s;
            }
            catch (FormatException)
            {
                return;
            }
            printPreviewDialog1.ShowDialog();
        }

        private async void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select 客戶 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 in('出貨退出單', '出貨單') and 刪除='0' order by 客戶";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
            });
            DataView dv = dt.DefaultView;
            DataTable DT = dv.ToTable("客戶", true);
            comboBox1.Items.Clear();
            for (int i = 0; i < DT.Rows.Count; i++) comboBox1.Items.Add(DT.Rows[i]["客戶"].ToString());
        }

        private async void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select 客戶 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 in('出貨退出單', '出貨單') and 刪除='0' order by 客戶";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
            });
            DataView dv = dt.DefaultView;
            DataTable DT = dv.ToTable("客戶", true);
            comboBox1.Items.Clear();
            for (int i = 0; i < DT.Rows.Count; i++) comboBox1.Items.Add(DT.Rows[i]["客戶"].ToString());
        }

        private async void comboBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            double total = 0.0;
            SQLConnect con = new SQLConnect();
            string SQL = "select 單子, 日期, 單子編號, 總金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 in('出貨退出單', '出貨單')" +
                         "and 客戶=@customer and 刪除='0' order by 單子";
            DataTable DT = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day and 單子 ='出貨退出單' " +
                    "and 客戶=@customer and 刪除='0'";
            DataTable dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            if (dt != null && dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["金額"].ToString() != "")
                {
                    total = double.Parse(dt.Rows[0]["金額"].ToString()) * -1.0;
                }
            }
            SQL = "select sum(總金額) as 金額 from 總單子_客戶 where 日期 between @first_day and @second_day " +
                    "and 單子 ='出貨單' and 客戶=@customer and 刪除='0'";
            dt = await con.searchDataTable(SQL, new
            {
                first_day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                second_day = dateTimePicker2.Value.ToString("yyyyMMdd"),
                customer = comboBox1.Text.ToString()
            });
            if (dt != null && dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["金額"].ToString() != "")
                {
                    total += double.Parse(dt.Rows[0]["金額"].ToString());
                }
            }
            label6.Text = total.ToString();
            if (DT != null)
            {
                if (DT.Rows.Count > 0)
                {
                    dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = DT;
                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                }
            }
        }
    }
}
