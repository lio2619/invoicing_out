using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 管理客戶 : Form
    {
        private string oldCompanyName = string.Empty;
        public 管理客戶()
        {
            InitializeComponent();
        }

        public async void search(string company, string number)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select * from 客戶 where 公司全名=@company";
            DataTable dt = await con.searchDataTable(SQL, new { company = company });
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = company;
                    textBox2.Text = dt.Rows[0]["送貨地址"].ToString();
                    textBox3.Text = dt.Rows[0]["聯絡電話一"].ToString();
                    textBox4.Text = dt.Rows[0]["傳真號碼"].ToString();
                    label6.Text = number;
                    oldCompanyName = company;
                }
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select (count(公司編號) + 1) as 編號 from 客戶";
            DataTable DT = con.Find(SQL);
            SQL = "insert into 客戶 (公司編號 ,公司全名, 送貨地址, 聯絡電話一, 傳真號碼) values (@number, @company, @address, @phone, @fax)";
            bool ret = await con.execute(SQL, new
            {
                number = DT.Rows[0]["編號"].ToString(),
                company = textBox1.Text,
                address = textBox2.Text,
                phone = textBox3.Text,
                fax = textBox4.Text
            });
            if (ret)
            {
                MessageBox.Show("加入客戶成功", "成功");
                oldCompanyName = textBox1.Text;
            }
            else MessageBox.Show("加入客戶失敗", "失敗");
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "update 客戶 set 送貨地址=@address, 聯絡電話一=@phone, 傳真號碼=@fax, 公司全名=@company where 公司全名=@number";
            bool ret = await con.execute(SQL, new
            {
                address = textBox2.Text,
                phone = textBox3.Text,
                fax = textBox4.Text,
                company = textBox1.Text,
                number = oldCompanyName
            });
            if (ret) MessageBox.Show("修改客戶成功", "成功");
            else MessageBox.Show("修改客戶失敗", "失敗");
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            string s = Interaction.InputBox("請輸入客戶名稱", "標題", "輸入框預設內容", -1, -1);
            SQLConnect con = new SQLConnect();
            string SQL = "select * from 客戶 where 公司全名=@company";
            DataTable dt = await con.searchDataTable(SQL, new { company = s });
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = s;
                    textBox2.Text = dt.Rows[0]["送貨地址"].ToString();
                    textBox3.Text = dt.Rows[0]["聯絡電話一"].ToString();
                    textBox4.Text = dt.Rows[0]["傳真號碼"].ToString();
                    oldCompanyName = s;
                }
                else MessageBox.Show("請輸入正確的客戶名稱", "錯誤");
            }
            else MessageBox.Show("請輸入正確的客戶名稱", "錯誤");
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            //string s = Interaction.InputBox("請輸入客戶名稱", "標題", "輸入框預設內容", -1, -1);
            string companyName = textBox1.Text;
            DialogResult result = MessageBox.Show($"確定要刪除{companyName}嗎？", "確認", MessageBoxButtons.YesNo);
            if(result == DialogResult.Yes)
            {
                SQLConnect con = new SQLConnect();
                string SQL = "delete from 客戶 where 公司全名=@company";
                bool ret = await con.execute(SQL, new { company = companyName });
                if (ret) MessageBox.Show("刪除客戶成功", "成功");
                else MessageBox.Show("刪除客戶失敗", "錯誤");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "0";
            textBox4.Text = "0";
        }
    }
}
