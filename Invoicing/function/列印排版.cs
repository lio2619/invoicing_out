using System.Drawing;
using System.Linq;

namespace Invoicing
{
    public class 列印排版
    {
        private int[] width = new int[] { 20, 135, 520, 575, 625, 680, 750 };
        private string[] name = new string[] { "進貨退出單", "進貨單" };
        public void print_header_setting(System.Drawing.Printing.PrintPageEventArgs e, string bili, string client, string date, string single_number,
                                                                                        string tex, string fax, string address, int page, int total_page)
        {
            if (name.Contains(bili))    //給廠商
            {
                e.Graphics.DrawString("", new Font("Arial", 22), Brushes.Black, new Point(20, 20));
                e.Graphics.DrawString(bili, new Font("Arial", 22), Brushes.Black, new Point(560, 20));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(20, 60));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(250, 60));
                e.Graphics.DrawString("頁次：" + (page + 1).ToString() + "/" + (total_page + 1).ToString(), new Font("Arial", 12), Brushes.Black, new Point(560, 60));
                e.Graphics.DrawString("廠商名稱：" + client, new Font("Arial", 12), Brushes.Black, new Point(20, 90));
                e.Graphics.DrawString("貨單日期：" + date, new Font("Arial", 12), Brushes.Black, new Point(560, 90));
                e.Graphics.DrawString("連絡電話：" + tex, new Font("Arial", 12), Brushes.Black, new Point(20, 120));
                e.Graphics.DrawString("傳真號碼：" + fax, new Font("Arial", 12), Brushes.Black, new Point(250, 120));
                e.Graphics.DrawString("貨單編號：" + single_number, new Font("Arial", 12), Brushes.Black, new Point(560, 120));
                e.Graphics.DrawString("送貨地址：" + address, new Font("Arial", 12), Brushes.Black, new Point(20, 150));
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 170);
                PointF point2 = new PointF(790, 170);
                e.Graphics.DrawLine(black, point1, point2);
                e.Graphics.DrawString("編號", new Font("Arial", 12), Brushes.Black, new Point(20, 170));
                e.Graphics.DrawString("品名", new Font("Arial", 12), Brushes.Black, new Point(140, 170));
                e.Graphics.DrawString("數量", new Font("Arial", 12), Brushes.Black, new Point(490, 170));
                e.Graphics.DrawString("單位", new Font("Arial", 12), Brushes.Black, new Point(540, 170));
                e.Graphics.DrawString("單價", new Font("Arial", 12), Brushes.Black, new Point(590, 170));
                e.Graphics.DrawString("金額", new Font("Arial", 12), Brushes.Black, new Point(640, 170));
                e.Graphics.DrawString("建議售價", new Font("Arial", 12), Brushes.Black, new Point(700, 170));
                PointF point3 = new PointF(20, 190);
                PointF point4 = new PointF(790, 190);
                e.Graphics.DrawLine(black, point3, point4);
            }
            else if (bili == "採購單")
            {
                e.Graphics.DrawString("", new Font("Arial", 22), Brushes.Black, new Point(20, 20));
                e.Graphics.DrawString(bili, new Font("Arial", 22), Brushes.Black, new Point(560, 20));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(20, 60));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(250, 60));
                e.Graphics.DrawString("頁次：" + (page + 1).ToString() + "/" + (total_page + 1).ToString(), new Font("Arial", 12), Brushes.Black, new Point(560, 60));
                e.Graphics.DrawString("廠商名稱：" + client, new Font("Arial", 12), Brushes.Black, new Point(20, 90));
                e.Graphics.DrawString("貨單日期：" + date, new Font("Arial", 12), Brushes.Black, new Point(560, 90));
                e.Graphics.DrawString("連絡電話：" + tex, new Font("Arial", 12), Brushes.Black, new Point(20, 120));
                e.Graphics.DrawString("傳真號碼：" + fax, new Font("Arial", 12), Brushes.Black, new Point(250, 120));
                e.Graphics.DrawString("貨單編號：" + single_number, new Font("Arial", 12), Brushes.Black, new Point(560, 120));
                e.Graphics.DrawString("送貨地址：" + address, new Font("Arial", 12), Brushes.Black, new Point(20, 150));
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 170);
                PointF point2 = new PointF(790, 170);
                e.Graphics.DrawLine(black, point1, point2);
                e.Graphics.DrawString("編號", new Font("Arial", 12), Brushes.Black, new Point(20, 170));
                e.Graphics.DrawString("品名", new Font("Arial", 12), Brushes.Black, new Point(140, 170));
                e.Graphics.DrawString("數量", new Font("Arial", 12), Brushes.Black, new Point(490, 170));
                e.Graphics.DrawString("單位", new Font("Arial", 12), Brushes.Black, new Point(540, 170));
                e.Graphics.DrawString("備註", new Font("Arial", 12), Brushes.Black, new Point(640, 170));
                PointF point3 = new PointF(20, 190);
                PointF point4 = new PointF(790, 190);
                e.Graphics.DrawLine(black, point3, point4);
            }
            else        //給客戶
            {
                e.Graphics.DrawString("", new Font("Arial", 22), Brushes.Black, new Point(20, 20));
                e.Graphics.DrawString(bili, new Font("Arial", 22), Brushes.Black, new Point(560, 20));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(20, 60));
                e.Graphics.DrawString("", new Font("Arial", 12), Brushes.Black, new Point(250, 60));
                e.Graphics.DrawString("頁次：" + (page + 1).ToString() + "/" + (total_page + 1).ToString(), new Font("Arial", 12), Brushes.Black, new Point(560, 60));
                e.Graphics.DrawString("客戶名稱：" + client, new Font("Arial", 12), Brushes.Black, new Point(20, 90));
                e.Graphics.DrawString("貨單日期：" + date, new Font("Arial", 12), Brushes.Black, new Point(560, 90));
                e.Graphics.DrawString("連絡電話：" + tex, new Font("Arial", 12), Brushes.Black, new Point(20, 120));
                e.Graphics.DrawString("傳真號碼：" + fax, new Font("Arial", 12), Brushes.Black, new Point(250, 120));
                e.Graphics.DrawString("貨單編號：" + single_number, new Font("Arial", 12), Brushes.Black, new Point(560, 120));
                e.Graphics.DrawString("送貨地址：" + address, new Font("Arial", 12), Brushes.Black, new Point(20, 150));
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 170);
                PointF point2 = new PointF(790, 170);
                e.Graphics.DrawLine(black, point1, point2);
                e.Graphics.DrawString("編號", new Font("Arial", 12), Brushes.Black, new Point(20, 170));
                e.Graphics.DrawString("品名", new Font("Arial", 12), Brushes.Black, new Point(140, 170));
                e.Graphics.DrawString("數量", new Font("Arial", 12), Brushes.Black, new Point(490, 170));
                e.Graphics.DrawString("單位", new Font("Arial", 12), Brushes.Black, new Point(540, 170));
                e.Graphics.DrawString("單價", new Font("Arial", 12), Brushes.Black, new Point(590, 170));
                e.Graphics.DrawString("金額", new Font("Arial", 12), Brushes.Black, new Point(640, 170));
                e.Graphics.DrawString("建議售價", new Font("Arial", 12), Brushes.Black, new Point(700, 170));
                PointF point3 = new PointF(20, 190);
                PointF point4 = new PointF(790, 190);
                e.Graphics.DrawLine(black, point3, point4);
            }
        }
        public void print_inside_setting(System.Drawing.Printing.PrintPageEventArgs e, string value, int high, int j, int row, string bili)
        {
            if (row % 5 == 0 && row != 0 && row % 40 != 0)
            {
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, high);
                PointF point2 = new PointF(790, high);
                e.Graphics.DrawLine(black, point1, point2);
            }
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Far;
            if (bili == "採購單")
            {
                if (j > 1)
                {
                    int[] width = new int[] { 20, 135, 520, 575, 680, 750 };
                    e.Graphics.DrawString(value, new Font("Arial", 10), Brushes.Black, new Point(width[j], high), stringFormat);
                }
                else
                {
                    e.Graphics.DrawString(value, new Font("Arial", 10), Brushes.Black, new Point(width[j], high));
                }
            }
            else
            {
                if (j > 1)
                {
                    e.Graphics.DrawString(value, new Font("Arial", 10), Brushes.Black, new Point(width[j], high), stringFormat);
                }
                else
                {
                    e.Graphics.DrawString(value, new Font("Arial", 10), Brushes.Black, new Point(width[j], high));
                }
            }

        }
        public void print_last_setting(System.Drawing.Printing.PrintPageEventArgs e, string value, string direction, int row, bool page)
        {
            if (page)
            {
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 1010);
                PointF point2 = new PointF(790, 1010);
                e.Graphics.DrawLine(black, point1, point2);
                e.Graphics.DrawString("備註：" + direction, new Font("Arial", 12), Brushes.Black, new Point(20, 1020));
                e.Graphics.DrawString("總計：" + value, new Font("Arial", 12), Brushes.Black, new Point(640, 1020));
            }
            else
            {
                if (row % 41 < 16)
                {
                    Pen black = new Pen(Color.Black);
                    PointF point1 = new PointF(20, 500);
                    PointF point2 = new PointF(790, 500);
                    e.Graphics.DrawLine(black, point1, point2);
                    e.Graphics.DrawString("備註：" + direction, new Font("Arial", 12), Brushes.Black, new Point(20, 510));
                    e.Graphics.DrawString("總計：" + value, new Font("Arial", 12), Brushes.Black, new Point(640, 510));
                }
                else
                {
                    Pen black = new Pen(Color.Black);
                    PointF point1 = new PointF(20, 1010);
                    PointF point2 = new PointF(790, 1010);
                    e.Graphics.DrawLine(black, point1, point2);
                    e.Graphics.DrawString("備註：" + direction, new Font("Arial", 12), Brushes.Black, new Point(20, 1010));
                    e.Graphics.DrawString("總計：" + value, new Font("Arial", 12), Brushes.Black, new Point(640, 1020));
                }
            }
        }
        public void collect_money_header(System.Drawing.Printing.PrintPageEventArgs e, string start_date, string end_date, string company)
        {
            e.Graphics.DrawString("", new Font("Arial", 20), Brushes.Black, new Point(320, 20));
            e.Graphics.DrawString("應收帳款簡要表", new Font("Arial", 16), Brushes.Black, new Point(325, 60));
            e.Graphics.DrawString("帳款區間：" + start_date + " ~ " + end_date, new Font("Arial", 12), Brushes.Black, new Point(20, 80));
            e.Graphics.DrawString("客戶名稱：" + company, new Font("Arial", 12), Brushes.Black, new Point(20, 110));
            e.Graphics.DrawString("單別", new Font("Arial", 12), Brushes.Black, new Point(70, 140));
            e.Graphics.DrawString("交易日期", new Font("Arial", 12), Brushes.Black, new Point(140, 140));
            e.Graphics.DrawString("交易單號", new Font("Arial", 12), Brushes.Black, new Point(260, 140));
            e.Graphics.DrawString("合計金額", new Font("Arial", 12), Brushes.Black, new Point(550, 140));
            e.Graphics.DrawString("總計金額", new Font("Arial", 12), Brushes.Black, new Point(650, 140));
            Pen black = new Pen(Color.Black);
            PointF point1 = new PointF(20, 160);
            PointF point2 = new PointF(790, 160);
            e.Graphics.DrawLine(black, point1, point2);
        }
        public void collect_money_inside(System.Drawing.Printing.PrintPageEventArgs e, string value, int high, int j)
        {
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Far;
            int[] Mwidth = new int[] { 120, 220, 365, 620, 720 };
            e.Graphics.DrawString(value, new Font("Arial", 12), Brushes.Black, new Point(Mwidth[j], high), stringFormat);
        }
        public void collect_money_last(System.Drawing.Printing.PrintPageEventArgs e, int row)
        {
            if (row % 41 < 16)
            {
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 500);
                PointF point2 = new PointF(790, 500);
                e.Graphics.DrawLine(black, point1, point2);
            }
            else
            {
                Pen black = new Pen(Color.Black);
                PointF point1 = new PointF(20, 1010);
                PointF point2 = new PointF(790, 1010);
                e.Graphics.DrawLine(black, point1, point2);
            }
        }
    }
}