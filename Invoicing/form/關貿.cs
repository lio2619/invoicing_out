using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using UglyToad.PdfPig;

namespace Invoicing.form
{
    public partial class 關貿 : Form
    {
        public 關貿()
        {
            InitializeComponent();
        }

        private void btnReadPdf_Click(object sender, EventArgs e)
        {
            DataTable DT = new DataTable();
            DT.Columns.Add("客戶名稱", typeof(string));
            DT.Columns.Add("關貿採購單號", typeof(string));
            DT.Columns.Add("進銷存單子編號", typeof(string));
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "pdf file (*.pdf)|*.pdf";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                List<(double MinX, double MaxX, bool CheckPattern)> xRanges = new List<(double, double, bool)>
                {
                    (43, 125, true),  //國碼
                    (331, 345, false), //數量
                    (410, 436, false)  //單價
                };
                double minY = 232;
                double maxY = 693;

                string file_path = dialog.FileName;
                string remarks = string.Empty;
                string name = string.Empty;
                string storeName = string.Empty;
                string number = string.Empty;
                string poNumber = string.Empty;
                List<string> ans = new List<string>();
                StringBuilder text = new StringBuilder();
                Regex regex = new Regex(@"^4\d{12}$");

                // 打開 PDF 文件
                using (PdfDocument document = PdfDocument.Open(file_path))
                {
                    // 遍歷每一頁
                    foreach (var page in document.GetPages())
                    {
                        var words = page.GetWords().ToList();
                        storeName = words[10].ToString();
                        number = words[18].ToString();
                        remarks = words[20].ToString();
                        var filteredWords = words
                                .Where(word =>
                                    xRanges.Any(range =>
                                        word.BoundingBox.Left >= range.MinX && word.BoundingBox.Right <= range.MaxX && // 篩選 X 軸範圍
                                        word.BoundingBox.Top >= minY && word.BoundingBox.Bottom <= maxY && // 篩選 Y 軸範圍
                                        (!range.CheckPattern || regex.IsMatch(word.Text)))) // 檢查是否符合數字模式（如果需要）
                                .Select(word => word.Text);
                        if (words[4].Text == words[6].Text)
                        {
                            ans.AddRange(filteredWords);
                            name = GetCustomerName(storeName.ToString(), number.ToString());
                            string num = TidyPdfListToDBData(ans, name, remarks);
                            DataRow newRow = DT.NewRow();
                            newRow["客戶名稱"] = name;
                            newRow["關貿採購單號"] = remarks;
                            newRow["進銷存單子編號"] = num;
                            DT.Rows.Add(newRow);
                            ans.Clear();
                        }
                        else
                        {
                            ans.AddRange(filteredWords);
                        }
                    }
                }
            }

            Datagridview(DT);
        }

        /// <summary>
        /// 將pdf讀取到的資料轉成可以儲存到DB的資料
        /// </summary>
        /// <param name="list"></param>
        private string TidyPdfListToDBData(List<string> list, string storeName, string poNumber)
        {
            SQLConnect con = new SQLConnect();
            string sql = "select count(單子編號) as 編號 from 總單子_客戶";
            DataTable num = con.Find(sql);
            var grouped = list.Select((value, index) => new { value, index })
                          .GroupBy(x => x.index / 3) // 每 3 個元素為一組
                          .Select(g => g.Select(x => x.value).ToList())
                          .Where(g => g.Count == 3) // 過濾掉元素數量小於 3 的組
                          .ToList();
            DataTable DBDT = new DataTable();
            DBDT.Columns.Add("單子編號", typeof(string));
            DBDT.Columns.Add("貨品編號", typeof(string));
            DBDT.Columns.Add("品名", typeof(string));
            DBDT.Columns.Add("數量", typeof(string));
            DBDT.Columns.Add("基本單位", typeof(string));
            DBDT.Columns.Add("單價", typeof(string));
            DBDT.Columns.Add("金額", typeof(string));
            DBDT.Columns.Add("備註", typeof(string));
            decimal totalCost = 0m;
            foreach (var group in grouped)
            {
                sql = $@"select 品名, 基本單位 from 貨品主檔 where 貨品編號 = '{group[2]}'";
                DataTable dt = con.Find(sql);
                DataRow newRow = DBDT.NewRow();
                decimal cost = decimal.Parse(group[0]) * decimal.Parse(group[1]);
                totalCost += cost;
                if(dt.Rows.Count == 0)
                {
                    string SQL = "insert into 貨品主檔 (貨品編號, 品名, 基本單位) values(@store," +
                        " @name, @unit)";
                    con.execute(SQL, new
                    {
                        store = group[2],
                        name = "空白",
                        unit = "空白"
                    });
                    newRow["單子編號"] = num.Rows[0]["編號"].ToString();
                    newRow["貨品編號"] = group[2];
                    newRow["品名"] = string.Empty;
                    newRow["數量"] = group[0];
                    newRow["基本單位"] = string.Empty;
                    newRow["單價"] = group[1];
                    newRow["金額"] = cost.ToString();
                    newRow["備註"] = string.Empty;
                }
                else
                {
                    newRow["單子編號"] = num.Rows[0]["編號"].ToString();
                    newRow["貨品編號"] = group[2];
                    newRow["品名"] = dt.Rows[0]["品名"].ToString();
                    newRow["數量"] = group[0];
                    newRow["基本單位"] = dt.Rows[0]["基本單位"].ToString();
                    newRow["單價"] = group[1];
                    newRow["金額"] = cost.ToString();
                    newRow["備註"] = string.Empty;
                }

                DBDT.Rows.Add(newRow);
            }
            //SavePdfFile(storeName, poNumber, totalCost.ToString(), num.Rows[0]["編號"].ToString());
            //SavePdfFileDetail(DBDT);

            return num.Rows[0]["編號"].ToString();
        }

        private void SavePdfFile(string storeName, string remark, string totalCost, string num)
        {
            SQLConnect con = new SQLConnect();
            string sql = "insert into 總單子_客戶 (日期, 時間, 客戶, 單子, 備註 ,總金額, 單子編號, 刪除, 關貿) values (@day, @time, @store_name, '出貨單'," +
                            " @remark, @total_cost, @number, '0', '1')";
            _ = con.execute(sql, new
            {
                day = DateTime.Now.ToString("yyyyMMdd"),
                time = DateTime.Now.ToString("yyyyMMddHHmmss"),
                store_name = storeName,
                remark = remark,
                total_cost = totalCost,
                number = num
            });
        }

        private void SavePdfFileDetail(DataTable dataTable)
        {
            SQLConnect con = new SQLConnect();
            _ = con.InsertList(dataTable);
        }

        private string GetCustomerName(string name, string number)
        {
            SQLConnect con = new SQLConnect();
            string sql = $@"select 公司全名 from 客戶 where 公司全名 like '%{name}({number.Substring(0, 4)}%'";
            DataTable dt = con.Find(sql);
            return dt.Rows[0]["公司全名"].ToString();
        }

        private void Datagridview(DataTable DT)
        {
            dgwConvertStore.AutoGenerateColumns = true;
            dgwConvertStore.DataSource = null;
            dgwConvertStore.Columns.Clear();
            dgwConvertStore.DataSource = DT;
            dgwConvertStore.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            // 禁用所有欄位的排序
            foreach (DataGridViewColumn column in dgwConvertStore.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dgwConvertStore.DefaultCellStyle.Font = new Font("Arial", 14);
        }
    }
}
