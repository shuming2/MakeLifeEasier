namespace MakeLifeEasier
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.BrowseInputButton = new System.Windows.Forms.Button();
            this.StartButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.InputExcelPathLabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // BrowseInputButton
            // 
            this.BrowseInputButton.Location = new System.Drawing.Point(271, 37);
            this.BrowseInputButton.Name = "BrowseInputButton";
            this.BrowseInputButton.Size = new System.Drawing.Size(75, 23);
            this.BrowseInputButton.TabIndex = 1;
            this.BrowseInputButton.Text = "浏览";
            this.BrowseInputButton.UseVisualStyleBackColor = true;
            this.BrowseInputButton.Click += new System.EventHandler(this.BrowseInputButton_Click);
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(271, 89);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(75, 23);
            this.StartButton.TabIndex = 3;
            this.StartButton.Text = "开始生成";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "导出文件:";
            // 
            // InputExcelPathLabel
            // 
            this.InputExcelPathLabel.AutoSize = true;
            this.InputExcelPathLabel.Location = new System.Drawing.Point(93, 42);
            this.InputExcelPathLabel.MaximumSize = new System.Drawing.Size(180, 20);
            this.InputExcelPathLabel.Name = "InputExcelPathLabel";
            this.InputExcelPathLabel.Size = new System.Drawing.Size(11, 12);
            this.InputExcelPathLabel.TabIndex = 6;
            this.InputExcelPathLabel.Text = "-";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "\"Excel文件|*.xls;*.xlsx\"";
            this.openFileDialog1.InitialDirectory = "Environment.GetFolderPath(Environment.SpecialFolder.Desktop";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(30, 91);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 21);
            this.dateTimePicker1.TabIndex = 8;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.Filter = "\"Excel文件|*.xls;*.xlsx\"";
            this.openFileDialog2.InitialDirectory = "Environment.GetFolderPath(Environment.SpecialFolder.Desktop";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(378, 136);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.InputExcelPathLabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartButton);
            this.Controls.Add(this.BrowseInputButton);
            this.Name = "Form1";
            this.Text = "MakeLifeEasier_v1.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button BrowseInputButton;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label InputExcelPathLabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
    }
}

