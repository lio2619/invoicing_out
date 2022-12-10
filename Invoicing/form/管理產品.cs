using Microsoft.VisualBasic;
using System;
using System.Data;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 管理產品 : Form
    {
        public 管理產品()
        {
            InitializeComponent();
        }

        public async void search(string number)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "select * from 貨品主檔 where 貨品編號=@store";
            DataTable dt = await con.searchDataTable(SQL, new {store = number});
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = number;
                    textBox2.Text = dt.Rows[0]["品名"].ToString();
                    textBox3.Text = dt.Rows[0]["基本單位"].ToString();
                    textBox4.Text = dt.Rows[0]["標準售價"].ToString();
                    textBox5.Text = dt.Rows[0]["售價A"].ToString();
                    textBox6.Text = dt.Rows[0]["售價B"].ToString();
                    textBox7.Text = dt.Rows[0]["售價C"].ToString();
                    textBox8.Text = dt.Rows[0]["標準成本"].ToString();
                    textBox9.Text = dt.Rows[0]["現行成本"].ToString();
                }
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "insert into 貨品主檔 (貨品編號, 品名, 基本單位, 標準售價, 售價A, 售價B, 售價C, 標準成本, 現行成本) values(@store," +
                        " @name, @unit, @price, @priceA, @priceB, @priceC, @cost, @now_cost)";
            bool ret = await con.execute(SQL, new
            {
                store = textBox1.Text,
                name = textBox2.Text,
                unit = textBox3.Text,
                price = textBox4.Text,
                priceA = textBox5.Text,
                priceB = textBox6.Text,
                priceC = textBox7.Text,
                cost = textBox8.Text,
                now_cost = textBox9.Text
            });
            if (ret) MessageBox.Show("加入產品成功", "成功");
            else MessageBox.Show("加入產品失敗", "失敗");
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            SQLConnect con = new SQLConnect();
            string SQL = "update 貨品主檔 set 品名=@name, 基本單位=@unit, 標準售價=@price, 售價A=@priceA, 售價B=@priceB, " +
                            "售價C=@priceC, 標準成本=@cost, 現行成本=@now_cost where 貨品編號=@store";
            bool ret = await con.execute(SQL, new
            {
                name = textBox2.Text,
                unit = textBox3.Text,
                price = textBox4.Text,
                priceA = textBox5.Text,
                priceB = textBox6.Text,
                priceC = textBox7.Text,
                cost = textBox8.Text,
                now_cost = textBox9.Text,
                store = textBox1.Text,
            });
            if (ret) MessageBox.Show("修改產品成功", "成功");
            else MessageBox.Show("修改產品失敗", "失敗");
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            string s = Interaction.InputBox("請輸入產品編號", "標題", "輸入框預設內容", -1, -1);
            SQLConnect con = new SQLConnect();
            string SQL = "select * from 貨品主檔 where 貨品編號=@store";
            DataTable dt = await con.searchDataTable(SQL, new {store = s});
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = s;
                    textBox2.Text = dt.Rows[0]["品名"].ToString();
                    textBox3.Text = dt.Rows[0]["基本單位"].ToString();
                    textBox4.Text = dt.Rows[0]["標準售價"].ToString();
                    textBox5.Text = dt.Rows[0]["售價A"].ToString();
                    textBox6.Text = dt.Rows[0]["售價B"].ToString();
                    textBox7.Text = dt.Rows[0]["售價C"].ToString();
                    textBox8.Text = dt.Rows[0]["標準成本"].ToString();
                    textBox9.Text = dt.Rows[0]["現行成本"].ToString();
                }
                else MessageBox.Show("請輸入正確的產品名稱", "錯誤");
            }
            else MessageBox.Show("請輸入正確的產品名稱", "錯誤");
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            string s = Interaction.InputBox("請輸入產品編號", "標題", "輸入框預設內容", -1, -1);
            SQLConnect con = new SQLConnect();
            string SQL = "delete from 貨品主檔 where 貨品編號=@store";
            bool ret = await con.execute(SQL, new { store = s });
            if (ret) MessageBox.Show("刪除產品成功", "成功");
            else MessageBox.Show("刪除產品失敗", "錯誤");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "0";
            textBox5.Text = "0";
            textBox6.Text = "0";
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBox9.Text = "0";
        }
    }
}
