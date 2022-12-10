using Invoicing.form;
using System;
using System.Windows.Forms;

namespace Invoicing
{
    public partial class 進銷存 : Form
    {
        public 進銷存()
        {
            InitializeComponent();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            採購單 form2 = new 採購單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 進貨退出單ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            進貨退出單 form2 = new 進貨退出單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 出貨單ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            出貨單 form2 = new 出貨單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 訂貨單ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            訂貨單 form2 = new 訂貨單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 進貨單ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            進貨單 form2 = new 進貨單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 出貨退出單ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            出貨退出單 form2 = new 出貨退出單();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 總進貨額ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            總進貨額 form2 = new 總進貨額();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 總金額ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            總金額 form2 = new 總金額();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 唐詣ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            唐詣 form2 = new 唐詣();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 個別資料ToolStripMenuItem_Click(object sender, EventArgs e)        //客戶
        {
            管理客戶 form2 = new 管理客戶();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 全部資料ToolStripMenuItem_Click(object sender, EventArgs e)        //客戶
        {
            所有客戶 form2 = new 所有客戶();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 個別資料ToolStripMenuItem1_Click(object sender, EventArgs e)       //商品
        {
            管理產品 form2 = new 管理產品();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 全部資料ToolStripMenuItem1_Click(object sender, EventArgs e)       //商品
        {
            所有商品 form2 = new 所有商品();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 個別資料ToolStripMenuItem2_Click(object sender, EventArgs e)       //廠商
        {
            管理廠商 form2 = new 管理廠商();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 全部資料ToolStripMenuItem2_Click(object sender, EventArgs e)       //廠商
        {
            所有廠商 form2 = new 所有廠商();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }

        private void 應收帳款ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            應收帳款 form2 = new 應收帳款();
            foreach (Form form in this.MdiChildren)
            {
                if (form.Name == form2.Name)
                {
                    form.Focus();
                    return;
                }
            }
            form2.MdiParent = this;
            form2.Show();
        }
    }
}
