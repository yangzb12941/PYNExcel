namespace PYNExcel
{
    partial class pynForm
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
            this.checkFileButton = new System.Windows.Forms.Button();
            this.fileNameTextBox = new System.Windows.Forms.TextBox();
            this.checkLabel = new System.Windows.Forms.Label();
            this.checkedSheetListBox = new System.Windows.Forms.CheckedListBox();
            this.checkedGoodsListBox = new System.Windows.Forms.CheckedListBox();
            this.goodsLabel = new System.Windows.Forms.Label();
            this.handleButton = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.sheetTrueButton = new System.Windows.Forms.Button();
            this.dataLabel = new System.Windows.Forms.Label();
            this.comName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.goodsName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ratioValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // checkFileButton
            // 
            this.checkFileButton.Location = new System.Drawing.Point(431, 31);
            this.checkFileButton.Name = "checkFileButton";
            this.checkFileButton.Size = new System.Drawing.Size(70, 21);
            this.checkFileButton.TabIndex = 0;
            this.checkFileButton.Text = "选择文件";
            this.checkFileButton.UseVisualStyleBackColor = true;
            this.checkFileButton.Click += new System.EventHandler(this.checkFileButton_Click);
            // 
            // fileNameTextBox
            // 
            this.fileNameTextBox.Location = new System.Drawing.Point(54, 31);
            this.fileNameTextBox.Name = "fileNameTextBox";
            this.fileNameTextBox.ReadOnly = true;
            this.fileNameTextBox.Size = new System.Drawing.Size(352, 21);
            this.fileNameTextBox.TabIndex = 1;
            // 
            // checkLabel
            // 
            this.checkLabel.AutoSize = true;
            this.checkLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkLabel.Location = new System.Drawing.Point(52, 65);
            this.checkLabel.Name = "checkLabel";
            this.checkLabel.Size = new System.Drawing.Size(122, 14);
            this.checkLabel.TabIndex = 2;
            this.checkLabel.Text = "请勾数据sheet页";
            // 
            // checkedSheetListBox
            // 
            this.checkedSheetListBox.CheckOnClick = true;
            this.checkedSheetListBox.FormattingEnabled = true;
            this.checkedSheetListBox.Location = new System.Drawing.Point(54, 82);
            this.checkedSheetListBox.Name = "checkedSheetListBox";
            this.checkedSheetListBox.Size = new System.Drawing.Size(162, 228);
            this.checkedSheetListBox.TabIndex = 3;
            // 
            // checkedGoodsListBox
            // 
            this.checkedGoodsListBox.CheckOnClick = true;
            this.checkedGoodsListBox.FormattingEnabled = true;
            this.checkedGoodsListBox.Location = new System.Drawing.Point(305, 82);
            this.checkedGoodsListBox.Name = "checkedGoodsListBox";
            this.checkedGoodsListBox.Size = new System.Drawing.Size(196, 228);
            this.checkedGoodsListBox.TabIndex = 5;
            this.checkedGoodsListBox.SelectedIndexChanged += new System.EventHandler(this.checkedGoodsListBox_SelectedIndexChanged);
            // 
            // goodsLabel
            // 
            this.goodsLabel.AutoSize = true;
            this.goodsLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.goodsLabel.Location = new System.Drawing.Point(303, 65);
            this.goodsLabel.Name = "goodsLabel";
            this.goodsLabel.Size = new System.Drawing.Size(82, 14);
            this.goodsLabel.TabIndex = 4;
            this.goodsLabel.Text = "请勾选物料";
            // 
            // handleButton
            // 
            this.handleButton.Location = new System.Drawing.Point(410, 511);
            this.handleButton.Name = "handleButton";
            this.handleButton.Size = new System.Drawing.Size(91, 27);
            this.handleButton.TabIndex = 6;
            this.handleButton.Text = "开始处理";
            this.handleButton.UseVisualStyleBackColor = true;
            this.handleButton.Click += new System.EventHandler(this.handleButton_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToResizeRows = false;
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView.CausesValidation = false;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.comName,
            this.goodsName,
            this.ratioValue});
            this.dataGridView.Location = new System.Drawing.Point(54, 340);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridView.RowTemplate.Height = 23;
            this.dataGridView.Size = new System.Drawing.Size(447, 160);
            this.dataGridView.TabIndex = 7;
            this.dataGridView.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEndEdit);
            // 
            // sheetTrueButton
            // 
            this.sheetTrueButton.Location = new System.Drawing.Point(224, 180);
            this.sheetTrueButton.Name = "sheetTrueButton";
            this.sheetTrueButton.Size = new System.Drawing.Size(75, 23);
            this.sheetTrueButton.TabIndex = 8;
            this.sheetTrueButton.Text = ">>>";
            this.sheetTrueButton.UseVisualStyleBackColor = true;
            this.sheetTrueButton.Click += new System.EventHandler(this.sheetTrueButton_Click);
            // 
            // dataLabel
            // 
            this.dataLabel.AutoSize = true;
            this.dataLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dataLabel.Location = new System.Drawing.Point(52, 320);
            this.dataLabel.Name = "dataLabel";
            this.dataLabel.Size = new System.Drawing.Size(82, 14);
            this.dataLabel.TabIndex = 10;
            this.dataLabel.Text = "公司补贴率";
            // 
            // comName
            // 
            this.comName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.comName.HeaderText = "公司抬头";
            this.comName.Name = "comName";
            this.comName.ReadOnly = true;
            this.comName.Width = 78;
            // 
            // goodsName
            // 
            this.goodsName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.goodsName.HeaderText = "品种";
            this.goodsName.Name = "goodsName";
            this.goodsName.ReadOnly = true;
            this.goodsName.Width = 54;
            // 
            // ratioValue
            // 
            this.ratioValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ratioValue.HeaderText = "补贴率";
            this.ratioValue.Name = "ratioValue";
            this.ratioValue.Width = 66;
            // 
            // pynForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 544);
            this.Controls.Add(this.dataLabel);
            this.Controls.Add(this.sheetTrueButton);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.handleButton);
            this.Controls.Add(this.checkedGoodsListBox);
            this.Controls.Add(this.goodsLabel);
            this.Controls.Add(this.checkedSheetListBox);
            this.Controls.Add(this.checkLabel);
            this.Controls.Add(this.fileNameTextBox);
            this.Controls.Add(this.checkFileButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "pynForm";
            this.Text = "浙江珏宏保税清关业务效率";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button checkFileButton;
        private System.Windows.Forms.TextBox fileNameTextBox;
        private System.Windows.Forms.Label checkLabel;
        private System.Windows.Forms.CheckedListBox checkedSheetListBox;
        private System.Windows.Forms.CheckedListBox checkedGoodsListBox;
        private System.Windows.Forms.Label goodsLabel;
        private System.Windows.Forms.Button handleButton;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button sheetTrueButton;
        private System.Windows.Forms.Label dataLabel;
        private System.Windows.Forms.DataGridViewTextBoxColumn comName;
        private System.Windows.Forms.DataGridViewTextBoxColumn goodsName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ratioValue;
    }
}

