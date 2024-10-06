using Invoicing.form;
using System;
using System.IO;
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

        private void 備份ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 取得當前目錄的上一層目錄
            string parentDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;

            // 指定要查找的檔案名稱
            string targetFileName = "Data.mdb";

            // 在上一層目錄中查找檔案
            string targetFilePath = Path.Combine(parentDirectory, targetFileName);

            if (!File.Exists(targetFilePath))
            {
                // 如果檔案不存在，讓使用者選擇檔案
                Console.WriteLine("檔案不存在，請選擇檔案。");
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        targetFilePath = openFileDialog.FileName;
                    }
                }
            }

            // 讓使用者選擇目標資料夾
            string targetFolderPath = string.Empty;
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "選擇要將檔案複製到的資料夾";

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    targetFolderPath = folderBrowserDialog.SelectedPath;
                }
            }

            // 複製檔案到目標資料夾
            string destinationFilePath = Path.Combine(targetFolderPath, targetFileName);
            try
            {
                File.Copy(targetFilePath, destinationFilePath, true); // true 參數表示如果檔案已存在則覆蓋
                MessageBox.Show("檔案備份成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"檔案備份失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 還原ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 創建一個檔案對話框讓使用者選擇要複製的檔案
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇要複製的檔案";
                openFileDialog.Filter = "所有檔案 (*.*)|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 取得選擇的檔案路徑
                    string sourceFilePath = openFileDialog.FileName;

                    // 讓使用者選擇目標資料夾
                    using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                    {
                        folderBrowserDialog.Description = "選擇複製檔案的資料夾";

                        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                        {
                            // 取得目標資料夾路徑
                            string targetFolderPath = folderBrowserDialog.SelectedPath;

                            // 確定目標檔案的完整路徑
                            string targetFilePath = Path.Combine(targetFolderPath, Path.GetFileName(sourceFilePath));

                            try
                            {
                                // 複製檔案到目標資料夾
                                File.Copy(sourceFilePath, targetFilePath, true);
                                MessageBox.Show("檔案複製成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"檔案複製失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }

        private void 金大ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            金大 form2 = new 金大();
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
