namespace Invoicing.form
{
    partial class 關貿
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnReadPdf = new System.Windows.Forms.Button();
            this.dgwConvertStore = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgwConvertStore)).BeginInit();
            this.SuspendLayout();
            // 
            // btnReadPdf
            // 
            this.btnReadPdf.Location = new System.Drawing.Point(230, 433);
            this.btnReadPdf.Name = "btnReadPdf";
            this.btnReadPdf.Size = new System.Drawing.Size(336, 54);
            this.btnReadPdf.TabIndex = 0;
            this.btnReadPdf.Text = "讀取";
            this.btnReadPdf.UseVisualStyleBackColor = true;
            this.btnReadPdf.Click += new System.EventHandler(this.btnReadPdf_Click);
            // 
            // dgwConvertStore
            // 
            this.dgwConvertStore.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dgwConvertStore.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgwConvertStore.Location = new System.Drawing.Point(12, 12);
            this.dgwConvertStore.Name = "dgwConvertStore";
            this.dgwConvertStore.RowTemplate.Height = 24;
            this.dgwConvertStore.Size = new System.Drawing.Size(794, 402);
            this.dgwConvertStore.TabIndex = 1;
            // 
            // 關貿
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(821, 499);
            this.Controls.Add(this.dgwConvertStore);
            this.Controls.Add(this.btnReadPdf);
            this.Name = "關貿";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "關貿";
            ((System.ComponentModel.ISupportInitialize)(this.dgwConvertStore)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnReadPdf;
        private System.Windows.Forms.DataGridView dgwConvertStore;
    }
}