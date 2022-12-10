﻿using Microsoft.VisualBasic;
using OfficeOpenXml;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection.Emit;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 採購單 : Form
    {
        private int page = 0;   //目前得頁數
        private int row = 0;   //目前第幾row
        private bool save = false;
        public 採購單()
        {
            InitializeComponent();
            datagridview();
            addbutton();
            label6.Visible = false;
            label7.Visible = false;
            SQLConnect con = new SQLConnect();
            string SQL = "select 公司全名 from 廠商 order by 公司全名";
            DataTable dt = con.Find(SQL);
            for (int i = 0; i < dt.Rows.Count; i++) comboBox1.Items.Add(dt.Rows[i]["公司全名"].ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
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

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                                                   dataGridView1.RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dataGridView1.RowHeadersDefaultCellStyle.Font,
                                   rectangle, dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private async void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            SQLConnect con = new SQLConnect();
            if (e.ColumnIndex == 0)
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[0].Value != null)
                {
                    string find = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    Form[] aryf = Parent.FindForm().MdiChildren;
                    foreach (Form f in aryf)
                    {
                        if (f.Name == "管理產品")
                        {
                            ((管理產品)f).search(find);
                            break;
                        }
                    }
                    string sql = "select 品名, 基本單位 from 貨品主檔 where 貨品編號=@find";
                    DataTable dt = await con.searchDataTable(sql, new { find = find });
                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = dt.Rows[0]["品名"].ToString();
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = dt.Rows[0]["基本單位"].ToString();
                        }
                    }
                    else MessageBox.Show("請輸入正確的編號", "錯誤");
                }

            }
        }

        private async void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            列印排版 form = new 列印排版();
            SQLConnect con = new SQLConnect();
            int total_page = (dataGridView1.RowCount - 1) / 2;
            int print_page = (dataGridView1.RowCount - 1) / 40;
            string store_name = comboBox1.Text.ToString();
            try
            {
                string SQL = "select 聯絡電話一, 傳真號碼, 送貨地址 from 廠商 where 公司全名=@store_name";
                DataTable dt = await con.searchDataTable(SQL, new { store_name = store_name });
                SQL = "select 單子編號 from 總單子_客戶 where 日期=@time and 客戶=@store_name and 單子='採購單' " +
                        "and 備註=@remark and 總金額='0'";
                DataTable dt_1 = await con.searchDataTable(SQL,
                    new
                    {
                        time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                        store_name = store_name,
                        remark = textBox1.Text.ToString(),
                    });
                int number = int.Parse(dt_1.Rows[0]["單子編號"].ToString());
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        form.print_header_setting(e, "採購單", store_name, dateTimePicker1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd") + number.ToString("000"),
                                        dt.Rows[0]["聯絡電話一"].ToString(), dt.Rows[0]["傳真號碼"].ToString(),
                                        dt.Rows[0]["送貨地址"].ToString(), page, print_page);
                    }
                    else MessageBox.Show("請輸入正確的廠商名稱", "錯誤");
                }
                else MessageBox.Show("請輸入正確的廠商名稱", "錯誤");
            }
            catch (IndexOutOfRangeException)
            {
                MessageBox.Show("請先儲存檔案");
                return;
            }
            int high = 195;     //內容起始高度
            for (; row < dataGridView1.RowCount - 1;)    //40就要換頁
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    form.print_inside_setting(e, dataGridView1.Rows[row].Cells[j].Value?.ToString() ?? "", high, j, row, "採購單");
                }
                if (row % 40 == 39 && row + 1 < dataGridView1.RowCount - 1)
                {
                    page++;
                    row++;
                    e.HasMorePages = (page < total_page);
                    form.print_last_setting(e, "", textBox1.Text, row, true);
                    return;
                }
                row++;
                high += 20;
            }
            form.print_last_setting(e, "", textBox1.Text, row, false);
            row = 0;
            page = 0;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("請選擇廠商");
                return;
            }
            SQLConnect con = new SQLConnect();
            string sql = "";

            if (label6.Visible || save)
            {
                sql = "delete from 整張儲存 where 單子編號=@number";
                await con.execute(sql, new { number = label6.Text.ToString() });
                read_dateGridView(label6.Text.ToString());
                sql = "update 總單子_客戶 set 總金額='0', 備註=@remark, 客戶=@customer, 日期=@time where 單子編號=@number";
                await con.execute(sql, new
                {
                    remark = textBox1.Text.ToString(),
                    customer = comboBox1.Text.ToString(),
                    time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    number = label6.Text.ToString()
                });
            }
            else
            {
                save = true;
                sql = "select count(單子編號) as 編號 from 總單子_客戶";
                DataTable num = con.Find(sql);
                if (num != null && num.Rows.Count > 0)
                {
                    sql = "insert into 總單子_客戶 (日期, 客戶, 單子, 備註 ,總金額, 單子編號) values (@time, @store_name, '採購單'," +
                            " @remark, '0', @number)";
                    await con.execute(sql, new
                    {
                        time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                        store_name = comboBox1.Text.ToString(),
                        remark = textBox1.Text.ToString(),
                        number = num.Rows[0]["編號"].ToString()
                    });
                    read_dateGridView(num.Rows[0]["編號"].ToString());
                }
                else
                {
                    sql = "insert into 總單子_客戶 (日期, 客戶, 單子, 備註 ,總金額, 單子編號) values (@time, @store_name, '進貨退出單'," +
                            " @remark, '0', 0)";
                    await con.execute(sql, new
                    {
                        time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                        store_name = comboBox1.Text.ToString(),
                        remark = textBox1.Text.ToString(),
                    });
                    read_dateGridView("0");
                }
            }
            MessageBox.Show("儲存完成");
        }
        private async void create_excel(string date, string file_name, string number)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select 聯絡電話一, 傳真號碼 from 廠商 where 公司全名=@store_name";
            int row = 5;
            DataTable dt = await con.searchDataTable(SQL, new { store_name = comboBox1.Text });
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"../單子/" + date + "/採購單/" + file_name + ".xlsx"); // 檔案路徑
            using (var excel = new ExcelPackage())
            {
                try
                {
                    var ws = excel.Workbook.Worksheets.Add("MySheet");

                    ws.Cells[1, 1].Value = "採購單";
                    ws.Cells[1, 7].Value = "貨單日期：" + dateTimePicker1.Value.ToString("yyyyMMdd");
                    ws.Cells[2, 1].Value = "廠商名稱：" + comboBox1.Text;
                    ws.Cells[2, 7].Value = "貨單編號：" + dateTimePicker1.Value.ToString("yyyyMMdd") + number;
                    ws.Cells[3, 1].Value = "連絡電話：" + dt.Rows[0]["聯絡電話一"].ToString();
                    ws.Cells[3, 4].Value = "傳真號碼：" + dt.Rows[0]["傳真號碼"].ToString();
                    ws.Cells[4, 1].Value = "編號";
                    ws.Cells[4, 2].Value = "品名";
                    ws.Cells[4, 5].Value = "數量";
                    ws.Cells[4, 6].Value = "單位";
                    ws.Cells[4, 7].Value = "備註";
                    int[] cell = new int[] { 1, 2, 5, 6, 7, 8, 9 };
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        {
                            ws.Cells[row, cell[j]].Value = dataGridView1.Rows[i].Cells[j].Value?.ToString() ?? "";
                        }
                        row++;
                    }
                    ws.Cells[row, 1].Value = "備註：" + textBox1.Text.ToString();
                    excel.SaveAs(file);
                }
                catch (InvalidOperationException)
                {
                    MessageBox.Show("已經有這個檔案了");
                    return;
                }
            }
            MessageBox.Show("完成");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult Result = MessageBox.Show("確定刷新", "詢問", MessageBoxButtons.YesNo);
            if (Result == DialogResult.Yes)
            {
                save = false;
                label6.Visible = false;
                label7.Visible = false;
                comboBox1.Text = "";
                datagridview();
                addbutton();
                textBox1.Text = "";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            讀取單子 form2 = new 讀取單子("採購單");
            label6.Visible = true;
            label7.Visible = true;
            form2.Owner = this;
            form2.Show();
        }
        private void datagridview()
        {
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            DataTable DT = new DataTable();
            DT.Columns.Add("貨品編號", typeof(string));
            DT.Columns.Add("品名", typeof(string));
            DT.Columns.Add("數量", typeof(string));
            DT.Columns.Add("基本單位", typeof(string));
            DT.Columns.Add("備註", typeof(string));
            dataGridView1.DataSource = DT;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoGenerateColumns = false;
        }
        private void addbutton()
        {
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            btn.Name = "search";
            btn.HeaderText = "查詢";
            btn.DefaultCellStyle.NullValue = "查詢";
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns.Add(btn);
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "search" && e.RowIndex >= 0)
            {
                模糊查詢 form2 = new 模糊查詢(e.RowIndex, "採購單");
                form2.Owner = this;
                form2.Show();
            }
        }
        private async Task read_dateGridView(string number)
        {
            DataTable DT = (DataTable)dataGridView1.DataSource;
            DT.AcceptChanges();
            DataColumn column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.Caption = "單子編號";
            column.ColumnName = "單子編號";
            column.DefaultValue = number;
            DT.Columns.Add(column);
            SQLConnect con = new SQLConnect();
            await con.InsertPurchase(DT);
            DT.Columns.Remove("單子編號");
            label6.Visible = true;
            label7.Visible = true;
            label6.Text = number;
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString("yyyyMM");
            string path = "../單子/" + date + "/採購單/";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            SQLConnect con = new SQLConnect();
            string SQL = "select 單子編號 from 總單子_客戶 where 日期=@time and 客戶=@store_name and 單子='採購單' " +
                        "and 備註=@remark and 總金額='0'";
            DataTable dt_1 = await con.searchDataTable(SQL,
                new
                {
                    time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    store_name = comboBox1.Text.ToString(),
                    remark = textBox1.Text.ToString()
                });
            int number = int.Parse(dt_1.Rows[0]["單子編號"].ToString());
            string file_name = "採購單_" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "_" + comboBox1.Text.ToString();
            create_excel(date, file_name, number.ToString("000"));
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            DataGridView dgv = sender as DataGridView;
            DialogResult dialog = MessageBox.Show("是否刪除", "詢問", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes)
            {
                int row_index = dgv.CurrentRow.Index;
                if (row_index < 0) return;
                try
                {
                    dgv.Rows.RemoveAt(row_index);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show("請先選擇要刪除的東西");
                    return;
                }
                catch (NullReferenceException ex)
                {
                    MessageBox.Show("沒有東西可以刪除");
                    return;
                }
            }
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
