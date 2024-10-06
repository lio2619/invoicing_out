using Microsoft.VisualBasic;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Invoicing.form
{
    public partial class 金大 : Form
    {
        private int page = 0;   //目前得頁數
        private int row = 0;   //目前第幾row
        private bool save = false;
        private DataTable source = new DataTable();

        public 金大()
        {
            InitializeComponent();
            label6.Visible = false;
            label7.Visible = false;
            SQLConnect con = new SQLConnect();
            string SQL = "select 公司全名 from 客戶 order by 公司全名";
            DataTable dt = con.Find(SQL);
            for (int i = 0; i < dt.Rows.Count; i++) comboBox1.Items.Add(dt.Rows[i]["公司全名"].ToString());
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                                                   dataGridView1.RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dataGridView1.RowHeadersDefaultCellStyle.Font,
                                   rectangle, dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private DataTable source_column(DataTable DT)
        {
            string[] s = new string[7] { "貨品編號", "品名", "基本單位", "數量", "單價", "金額", "備註" };
            foreach (string ss in s)
            {
                DataColumn column = new DataColumn();
                column.DataType = Type.GetType("System.String");
                column.Caption = ss;
                column.ColumnName = ss;
                DT.Columns.Add(column);
            }
            return DT;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable DT = new DataTable();
            DataTable now_source = source_column(source);
            string sql = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "excel file (*.xls)|*.xls";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string file_path = dialog.FileName;
                string provider = "Microsoft.Jet.OLEDB.4.0;";
                string extendedstring = "'Excel 8.0;";
                string connect_string = "Data Source=" + file_path + ";Provider=" + provider + "Extended Properties=" + extendedstring + "HDR=No';";

                using (OleDbConnection connect = new OleDbConnection(connect_string))
                {
                    connect.Open();
                    DataTable dtExcelSchema = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    //SheetName = "復原_工作表1$";
                    sql = "select F1, F3, F6, F7, F8 from ["+ SheetName + "A16:K] where IsNumeric(F1) = True";
                    using (OleDbDataAdapter dr = new OleDbDataAdapter(sql, connect))
                    {
                        dr.Fill(DT);
                    }
                    connect.Close();
                }
            }
            SQLConnect con = new SQLConnect();
            dataGridView1.Rows.Add(DT.Rows.Count);
            double total = 0.0;
            string cost = "";
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                if (DT.Rows[i]["F1"].ToString() == "") continue;
                sql = "select 貨品編號 as F1, 品名, 基本單位, 標準售價,標準成本 from 貨品主檔 where 貨品編號='" + DT.Rows[i]["F1"].ToString() + "'";
                DataTable grid = con.Find(sql);
                if (grid != null)
                {
                    if (grid.Rows.Count > 0)
                    {
                        cost = (double.Parse(grid.Rows[0]["標準售價"].ToString()) * double.Parse(DT.Rows[i]["F6"].ToString())).ToString();
                        /*if (!decimal.Equals((decimal.Parse(DT.Rows[i]["F10"].ToString())) / 1.05m, decimal.Parse(grid.Rows[0]["標準成本"].ToString())))
                        {
                            MessageBox.Show($"品名：{grid.Rows[0]["品名"]} 漲價");
                        }*/
                    }
                    else if (string.IsNullOrEmpty(DT.Rows[i]["F3"].ToString()))
                    {
                        break;
                    }
                    else
                    {
                        cost = "0";
                    }
                }
                DataRow workRow = now_source.NewRow();
                dataGridView1.Rows[i].Cells[0].Value = DT.Rows[i]["F1"].ToString();      //編號
                workRow["貨品編號"] = DT.Rows[i]["F1"].ToString();
                if (grid.Rows.Count == 0)
                {
                    dataGridView1.Rows[i].Cells[1].Value = DT.Rows[i]["F3"].ToString();
                    dataGridView1.Rows[i].Cells[2].Value = DT.Rows[i]["F6"].ToString();
                    dataGridView1.Rows[i].Cells[3].Value = DT.Rows[i]["F7"].ToString();
                    workRow["品名"] = DT.Rows[i]["F3"].ToString();
                    workRow["基本單位"] = DT.Rows[i]["F7"].ToString();
                    workRow["數量"] = DT.Rows[i]["F6"].ToString();
                    string SQL = "insert into 貨品主檔 (貨品編號, 品名, 基本單位) values(@store," +
                        " @name, @unit)";
                    con.execute(SQL, new
                    {
                        store = DT.Rows[i]["F1"].ToString(),
                        name = DT.Rows[i]["F3"].ToString(),
                        unit = DT.Rows[i]["F7"].ToString()
                    });
                }
                else
                {
                    sql = "select 建議售價 from 建議售價 where 標準售價='" + grid.Rows[0]["標準售價"].ToString() + "'";
                    DataTable sug = con.Find(sql);
                    dataGridView1.Rows[i].Cells[1].Value = grid.Rows[0]["品名"].ToString();
                    dataGridView1.Rows[i].Cells[2].Value = DT.Rows[i]["F6"].ToString();
                    dataGridView1.Rows[i].Cells[3].Value = grid.Rows[0]["基本單位"].ToString();
                    dataGridView1.Rows[i].Cells[4].Value = grid.Rows[0]["標準售價"].ToString();
                    dataGridView1.Rows[i].Cells[5].Value = cost;
                    try
                    {
                        dataGridView1.Rows[i].Cells[6].Value = sug.Rows[0]["建議售價"].ToString();
                        workRow["備註"] = sug.Rows[0]["建議售價"].ToString();
                    }
                    catch
                    {
                        MessageBox.Show($"此單價：{grid.Rows[0]["標準售價"]}沒有建議售價");
                    }
                    workRow["品名"] = grid.Rows[0]["品名"].ToString();
                    workRow["基本單位"] = grid.Rows[0]["基本單位"].ToString();
                    workRow["數量"] = DT.Rows[i]["F6"].ToString();
                    workRow["單價"] = grid.Rows[0]["標準售價"].ToString();
                    workRow["金額"] = cost;
                }
                now_source.Rows.Add(workRow);
                total += double.Parse(cost);
            }
            label2.Text = total.ToString();
            source = now_source;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("請選擇客戶");
                return;
            }
            if (save)
            {
                MessageBox.Show("已經儲存過");
                return;
            }
            SQLConnect con = new SQLConnect();
            string sql = "";
            sql = "select count(單子編號) as 編號 from 總單子_客戶";
            DataTable num = con.Find(sql);
            if (num != null && num.Rows.Count > 0)
            {
                sql = "insert into 總單子_客戶 (日期, 時間, 客戶, 單子, 備註 ,總金額, 單子編號, 刪除) values (@day, @time, @store_name, '出貨單'," +
                            " @remark, @total_cost, @number, '0')";
                await con.execute(sql, new
                {
                    day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    time = DateTime.Now.ToString("yyyyMMddHHmmss"),
                    store_name = comboBox1.Text.ToString(),
                    remark = textBox1.Text.ToString(),
                    total_cost = label2.Text.ToString(),
                    number = num.Rows[0]["編號"].ToString()
                });
                read_dateGridView(num.Rows[0]["編號"].ToString());
            }
            else
            {
                sql = "insert into 總單子_客戶 (日期, 時間, 客戶, 單子, 備註 ,總金額, 單子編號, 刪除) values (@day, @time, @store_name, '出貨單'," +
                            " @remark, @total_cost, '0', '0')";
                await con.execute(sql, new
                {
                    day = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    time = DateTime.Now.ToString("yyyyMMddHHmmss"),
                    store_name = comboBox1.Text.ToString(),
                    remark = textBox1.Text.ToString(),
                    total_cost = label2.Text.ToString()
                });
                read_dateGridView("0");
            }
            save = true;
            MessageBox.Show("儲存完成");
        }

        private async Task read_dateGridView(string number)
        {
            DataColumn column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.Caption = "單子編號";
            column.ColumnName = "單子編號";
            column.DefaultValue = number;
            source.Columns.Add(column);
            SQLConnect con = new SQLConnect();
            await con.InsertList(source);
            source.Columns.Remove("單子編號");
            label6.Visible = true;
            label7.Visible = true;
            label6.Text = number;
        }

        private void button3_Click(object sender, EventArgs e)
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

        private async void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            列印排版 form = new 列印排版();
            SQLConnect con = new SQLConnect();
            int total_page = (dataGridView1.RowCount - 1) / 2;
            int print_page = (dataGridView1.RowCount - 1) / 40;
            int number = 0;
            string store_name = this.comboBox1.Text.ToString();
            string SQL = "select 聯絡電話一, 傳真號碼, 送貨地址 from 客戶 where 公司全名=@store_name";
            DataTable dt = await con.searchDataTable(SQL, new { store_name = store_name });
            SQL = "select 單子編號 from 總單子_客戶 where 日期=@time and 客戶=@store_name and 單子='出貨單' " +
                    "and 備註=@remark and 總金額=@total_cost";
            DataTable dt_1 = await con.searchDataTable(SQL,
                new
                {
                    time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    store_name = store_name,
                    remark = textBox1.Text.ToString(),
                    total_cost = label2.Text.ToString()
                });
            try
            {
                number = int.Parse(dt_1.Rows[0]["單子編號"].ToString());
            }
            catch
            {
                MessageBox.Show("請先儲存檔案");
                return;
            }
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    form.print_header_setting(e, "出貨單", store_name, dateTimePicker1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd") + number.ToString("000"),
                    dt.Rows[0]["聯絡電話一"].ToString(), dt.Rows[0]["傳真號碼"].ToString(),
                                    dt.Rows[0]["送貨地址"].ToString(), page, print_page);
                }
                else MessageBox.Show("請輸入正確的客戶名稱", "錯誤");
            }
            else MessageBox.Show("請輸入正確的客戶名稱", "錯誤");
            int high = 195;     //內容起始高度
            for (; row < dataGridView1.RowCount - 1;)    //40就要換頁
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    form.print_inside_setting(e, dataGridView1.Rows[row].Cells[j].Value?.ToString() ?? "", high, j, row, "出貨單");
                }
                if (row % 40 == 39 && row + 1 < dataGridView1.RowCount - 1)
                {
                    page++;
                    row++;
                    e.HasMorePages = (page < total_page);
                    form.print_last_setting(e, this.label2.Text.ToString(), this.textBox1.Text, row, true);
                    return;
                }
                row++;
                high += 20;
            }
            form.print_last_setting(e, this.label2.Text.ToString(), this.textBox1.Text, row, false);
            row = 0;
            page = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult Result = MessageBox.Show("確定刷新", "詢問", MessageBoxButtons.YesNo);
            if (Result == DialogResult.Yes)
            {
                comboBox1.Text = "";
                dataGridView1.Rows.Clear();
                label2.Text = "0";
                textBox1.Text = "";
                save = false;
                source.Rows.Clear();
                source.Columns.Clear();
                label6.Visible = false;
                label7.Visible = false;
            }
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString("yyyyMM");
            string path = "../單子/" + date + "/出貨單/";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            SQLConnect con = new SQLConnect();
            string SQL = "select 單子編號 from 總單子_客戶 where 日期=@time and 客戶=@store_name and 單子='出貨單' " +
                        "and 備註=@remark and 總金額=@total_cost";
            DataTable dt_1 = await con.searchDataTable(SQL,
                new
                {
                    time = dateTimePicker1.Value.ToString("yyyyMMdd"),
                    store_name = comboBox1.Text.ToString(),
                    remark = textBox1.Text.ToString(),
                    total_cost = label2.Text.ToString()
                });
            int number = int.Parse(dt_1.Rows[0]["單子編號"].ToString());
            string file_name = "出貨單_" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "_" + comboBox1.Text.ToString();
            create_excel(date, file_name, number.ToString("000"));
        }

        private async void create_excel(string date, string file_name, string number)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select 聯絡電話一, 傳真號碼 from 客戶 where 公司全名=@store_name";
            int row = 5;
            var file = new FileInfo(file_name);
            DataTable dt = await con.searchDataTable(SQL, new { store_name = comboBox1.Text });
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            file = new FileInfo(@"../單子/" + date + "/出貨單/" + file_name.Split('(')[0] + "_" + number + ".xlsx");

            using (var excel = new ExcelPackage(file))
            {
                try
                {
                    ExcelWorksheet ws;
                    ws = excel.Workbook.Worksheets.Add("MySheet");

                    if (dt.Rows.Count > 0)
                    {
                        ws.Cells[1, 1].Value = "出貨單";
                        ws.Cells[1, 7].Value = "貨單日期：" + dateTimePicker1.Value.ToString("yyyyMMdd");
                        ws.Cells[2, 1].Value = "客戶名稱：";
                        ws.Cells[2, 2].Value = comboBox1.Text;
                        ws.Cells[2, 7].Value = "貨單編號：" + dateTimePicker1.Value.ToString("yyyyMMdd") + number;
                        ws.Cells[3, 1].Value = "連絡電話：" + dt.Rows[0]["聯絡電話一"].ToString();
                        ws.Cells[3, 4].Value = "傳真號碼：" + dt.Rows[0]["傳真號碼"].ToString();
                        ws.Cells[4, 1].Value = "編號";
                        ws.Cells[4, 2].Value = "品名";
                        ws.Cells[4, 5].Value = "數量";
                        ws.Cells[4, 6].Value = "單位";
                        ws.Cells[4, 7].Value = "單價";
                        ws.Cells[4, 8].Value = "金額";
                        ws.Cells[4, 9].Value = "備註";
                    }
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
                    ws.Cells[row, 8].Value = "總計：" + label2.Text.ToString();
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

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
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
                    double price = dataGridView1.Rows.Cast<DataGridViewRow>()
                                .Sum(r => Convert.ToDouble(r.Cells["金額"].Value));
                    label2.Text = price.ToString();
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
                catch (ArgumentException ex)
                {
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

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)     //把單價跟數量乘起來放到金額，並且把金額加起來放到label2
            {
                try
                {
                    double number = double.Parse(dataGridView1.Rows[e.RowIndex].Cells[4].Value?.ToString() ?? "0");    //數量，如果沒有輸入數字就填入0
                    double cost = double.Parse(dataGridView1.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "0");      //單價，如果沒有輸入數字就填入0
                    if (number <= 0 || cost <= 0)
                    {
                        MessageBox.Show("請輸入數字", "錯誤");
                    }
                    double a = double.Parse(dataGridView1.Rows[e.RowIndex].Cells[5].Value?.ToString() ?? "0");
                    dataGridView1.Rows[e.RowIndex].Cells[5].Value = number * cost;
                    double originPrice = double.Parse(label2.Text);
                    double price = number * cost + originPrice - a;
                    label2.Text = price.ToString();
                }
                catch (FormatException)
                {
                    MessageBox.Show("請輸入數量");
                    return;
                }
            }
        }
    }
}
