namespace SouthernCluster.ReportCards
{
    partial class ReportCardWindowsGUIMerger
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.clbNames = new System.Windows.Forms.CheckedListBox();
            this.grpNames = new System.Windows.Forms.GroupBox();
            this.cbSelectAll = new System.Windows.Forms.CheckBox();
            this.grpMergeType = new System.Windows.Forms.GroupBox();
            this.btnSavetoBrowse = new System.Windows.Forms.Button();
            this.lblSaveTo = new System.Windows.Forms.Label();
            this.cbMergeToPDF = new System.Windows.Forms.CheckBox();
            this.cbMergeToPUB = new System.Windows.Forms.CheckBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnMerge = new System.Windows.Forms.Button();
            this.ssMergeStatus = new System.Windows.Forms.StatusStrip();
            this.ssStatusText = new System.Windows.Forms.ToolStripStatusLabel();
            this.ssProgress = new System.Windows.Forms.ToolStripProgressBar();
            this.fdTemplateOpen = new System.Windows.Forms.OpenFileDialog();
            this.fdSaveTo = new System.Windows.Forms.FolderBrowserDialog();
            this.grpMergeFrom = new System.Windows.Forms.GroupBox();
            this.btnDatasourceBrowse = new System.Windows.Forms.Button();
            this.lblDatasource = new System.Windows.Forms.Label();
            this.btnTemplateBrowse = new System.Windows.Forms.Button();
            this.lblTemplate = new System.Windows.Forms.Label();
            this.fdDatasourceOpen = new System.Windows.Forms.OpenFileDialog();
            this.btnPrint = new System.Windows.Forms.Button();
            this.tbSaveTo = new System.Windows.Forms.TextBox();
            this.tbDatasource = new System.Windows.Forms.TextBox();
            this.tbTemplate = new System.Windows.Forms.TextBox();
            this.grpNames.SuspendLayout();
            this.grpMergeType.SuspendLayout();
            this.ssMergeStatus.SuspendLayout();
            this.grpMergeFrom.SuspendLayout();
            this.SuspendLayout();
            // 
            // clbNames
            // 
            this.clbNames.CheckOnClick = true;
            this.clbNames.FormattingEnabled = true;
            this.clbNames.Location = new System.Drawing.Point(6, 42);
            this.clbNames.Name = "clbNames";
            this.clbNames.Size = new System.Drawing.Size(150, 124);
            this.clbNames.TabIndex = 0;
            this.clbNames.ThreeDCheckBoxes = true;
            this.clbNames.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.clbNames_ItemCheck);
            // 
            // grpNames
            // 
            this.grpNames.Controls.Add(this.cbSelectAll);
            this.grpNames.Controls.Add(this.clbNames);
            this.grpNames.Location = new System.Drawing.Point(320, 12);
            this.grpNames.Name = "grpNames";
            this.grpNames.Size = new System.Drawing.Size(164, 179);
            this.grpNames.TabIndex = 2;
            this.grpNames.TabStop = false;
            this.grpNames.Text = "&Reports to Merge";
            // 
            // cbSelectAll
            // 
            this.cbSelectAll.AutoSize = true;
            this.cbSelectAll.Location = new System.Drawing.Point(9, 19);
            this.cbSelectAll.Name = "cbSelectAll";
            this.cbSelectAll.Size = new System.Drawing.Size(70, 17);
            this.cbSelectAll.TabIndex = 2;
            this.cbSelectAll.Text = "Select All";
            this.cbSelectAll.UseVisualStyleBackColor = true;
            this.cbSelectAll.CheckedChanged += new System.EventHandler(this.cbSelectAll_CheckedChanged);
            // 
            // grpMergeType
            // 
            this.grpMergeType.Controls.Add(this.btnSavetoBrowse);
            this.grpMergeType.Controls.Add(this.lblSaveTo);
            this.grpMergeType.Controls.Add(this.tbSaveTo);
            this.grpMergeType.Controls.Add(this.cbMergeToPDF);
            this.grpMergeType.Controls.Add(this.cbMergeToPUB);
            this.grpMergeType.Location = new System.Drawing.Point(12, 94);
            this.grpMergeType.Name = "grpMergeType";
            this.grpMergeType.Size = new System.Drawing.Size(302, 73);
            this.grpMergeType.TabIndex = 1;
            this.grpMergeType.TabStop = false;
            this.grpMergeType.Text = "Merge To ...";
            // 
            // btnSavetoBrowse
            // 
            this.btnSavetoBrowse.Location = new System.Drawing.Point(242, 40);
            this.btnSavetoBrowse.Name = "btnSavetoBrowse";
            this.btnSavetoBrowse.Size = new System.Drawing.Size(53, 23);
            this.btnSavetoBrowse.TabIndex = 4;
            this.btnSavetoBrowse.Text = "Browse";
            this.btnSavetoBrowse.UseVisualStyleBackColor = true;
            this.btnSavetoBrowse.Click += new System.EventHandler(this.btnSavetoBrowse_Click);
            // 
            // lblSaveTo
            // 
            this.lblSaveTo.AutoSize = true;
            this.lblSaveTo.Location = new System.Drawing.Point(6, 45);
            this.lblSaveTo.Name = "lblSaveTo";
            this.lblSaveTo.Size = new System.Drawing.Size(51, 13);
            this.lblSaveTo.TabIndex = 2;
            this.lblSaveTo.Text = "&Save To:";
            // 
            // cbMergeToPDF
            // 
            this.cbMergeToPDF.AutoSize = true;
            this.cbMergeToPDF.Location = new System.Drawing.Point(192, 19);
            this.cbMergeToPDF.Name = "cbMergeToPDF";
            this.cbMergeToPDF.Size = new System.Drawing.Size(104, 17);
            this.cbMergeToPDF.TabIndex = 1;
            this.cbMergeToPDF.Text = "PD&F Documents";
            this.cbMergeToPDF.UseVisualStyleBackColor = true;
            // 
            // cbMergeToPUB
            // 
            this.cbMergeToPUB.AutoSize = true;
            this.cbMergeToPUB.Location = new System.Drawing.Point(6, 19);
            this.cbMergeToPUB.Name = "cbMergeToPUB";
            this.cbMergeToPUB.Size = new System.Drawing.Size(126, 17);
            this.cbMergeToPUB.TabIndex = 0;
            this.cbMergeToPUB.Text = "&Publisher Documents";
            this.cbMergeToPUB.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(217, 173);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnMerge
            // 
            this.btnMerge.Enabled = false;
            this.btnMerge.Location = new System.Drawing.Point(18, 173);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new System.Drawing.Size(90, 23);
            this.btnMerge.TabIndex = 5;
            this.btnMerge.Text = "&Merge";
            this.btnMerge.UseVisualStyleBackColor = true;
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            // 
            // ssMergeStatus
            // 
            this.ssMergeStatus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ssStatusText,
            this.ssProgress});
            this.ssMergeStatus.Location = new System.Drawing.Point(0, 203);
            this.ssMergeStatus.Name = "ssMergeStatus";
            this.ssMergeStatus.Size = new System.Drawing.Size(495, 22);
            this.ssMergeStatus.SizingGrip = false;
            this.ssMergeStatus.TabIndex = 5;
            this.ssMergeStatus.Text = "Status";
            // 
            // ssStatusText
            // 
            this.ssStatusText.AutoSize = false;
            this.ssStatusText.Name = "ssStatusText";
            this.ssStatusText.Size = new System.Drawing.Size(350, 17);
            this.ssStatusText.Text = "Status";
            this.ssStatusText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ssProgress
            // 
            this.ssProgress.Name = "ssProgress";
            this.ssProgress.Size = new System.Drawing.Size(100, 16);
            // 
            // fdTemplateOpen
            // 
            this.fdTemplateOpen.DefaultExt = "pub";
            this.fdTemplateOpen.Filter = "Microsoft Publisher Documents|*.pub";
            this.fdTemplateOpen.ReadOnlyChecked = true;
            this.fdTemplateOpen.RestoreDirectory = true;
            // 
            // grpMergeFrom
            // 
            this.grpMergeFrom.Controls.Add(this.btnDatasourceBrowse);
            this.grpMergeFrom.Controls.Add(this.tbDatasource);
            this.grpMergeFrom.Controls.Add(this.lblDatasource);
            this.grpMergeFrom.Controls.Add(this.btnTemplateBrowse);
            this.grpMergeFrom.Controls.Add(this.tbTemplate);
            this.grpMergeFrom.Controls.Add(this.lblTemplate);
            this.grpMergeFrom.Location = new System.Drawing.Point(12, 12);
            this.grpMergeFrom.Name = "grpMergeFrom";
            this.grpMergeFrom.Size = new System.Drawing.Size(301, 76);
            this.grpMergeFrom.TabIndex = 0;
            this.grpMergeFrom.TabStop = false;
            this.grpMergeFrom.Text = "Merge From ...";
            // 
            // btnDatasourceBrowse
            // 
            this.btnDatasourceBrowse.Location = new System.Drawing.Point(242, 45);
            this.btnDatasourceBrowse.Name = "btnDatasourceBrowse";
            this.btnDatasourceBrowse.Size = new System.Drawing.Size(53, 23);
            this.btnDatasourceBrowse.TabIndex = 5;
            this.btnDatasourceBrowse.Text = "Browse";
            this.btnDatasourceBrowse.UseVisualStyleBackColor = true;
            this.btnDatasourceBrowse.Click += new System.EventHandler(this.btnDatasourceBrowse_Click);
            // 
            // lblDatasource
            // 
            this.lblDatasource.AutoSize = true;
            this.lblDatasource.Location = new System.Drawing.Point(6, 50);
            this.lblDatasource.Name = "lblDatasource";
            this.lblDatasource.Size = new System.Drawing.Size(65, 13);
            this.lblDatasource.TabIndex = 3;
            this.lblDatasource.Text = "&Datasource:";
            // 
            // btnTemplateBrowse
            // 
            this.btnTemplateBrowse.Location = new System.Drawing.Point(242, 17);
            this.btnTemplateBrowse.Name = "btnTemplateBrowse";
            this.btnTemplateBrowse.Size = new System.Drawing.Size(53, 23);
            this.btnTemplateBrowse.TabIndex = 2;
            this.btnTemplateBrowse.Text = "Browse";
            this.btnTemplateBrowse.UseVisualStyleBackColor = true;
            this.btnTemplateBrowse.Click += new System.EventHandler(this.btnTemplateBrowse_Click);
            // 
            // lblTemplate
            // 
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Location = new System.Drawing.Point(6, 22);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new System.Drawing.Size(54, 13);
            this.lblTemplate.TabIndex = 0;
            this.lblTemplate.Text = "&Template:";
            // 
            // fdDatasourceOpen
            // 
            this.fdDatasourceOpen.DefaultExt = "xlsx";
            this.fdDatasourceOpen.Filter = "Microsoft Excel Workbook|*.xlsx;*.xls;*.xlsb|Excel Xml Workbook|*.xlsx|Excel Bina" +
    "ry Workbook|*.xlsb|Excel 97-2003 Workbook|*.xls";
            this.fdDatasourceOpen.ReadOnlyChecked = true;
            this.fdDatasourceOpen.RestoreDirectory = true;
            // 
            // btnPrint
            // 
            this.btnPrint.Enabled = false;
            this.btnPrint.Location = new System.Drawing.Point(114, 173);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(90, 23);
            this.btnPrint.TabIndex = 7;
            this.btnPrint.Text = "Print";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // tbSaveTo
            // 
            this.tbSaveTo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tbSaveTo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories;
            this.tbSaveTo.Location = new System.Drawing.Point(77, 42);
            this.tbSaveTo.Name = "tbSaveTo";
            this.tbSaveTo.ReadOnly = true;
            this.tbSaveTo.Size = new System.Drawing.Size(159, 20);
            this.tbSaveTo.TabIndex = 3;
            // 
            // tbDatasource
            // 
            this.tbDatasource.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tbDatasource.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem;
            this.tbDatasource.Location = new System.Drawing.Point(77, 47);
            this.tbDatasource.Name = "tbDatasource";
            this.tbDatasource.ReadOnly = true;
            this.tbDatasource.Size = new System.Drawing.Size(159, 20);
            this.tbDatasource.TabIndex = 4;
            // 
            // tbTemplate
            // 
            this.tbTemplate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tbTemplate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem;
            this.tbTemplate.Location = new System.Drawing.Point(77, 19);
            this.tbTemplate.Name = "tbTemplate";
            this.tbTemplate.ReadOnly = true;
            this.tbTemplate.Size = new System.Drawing.Size(159, 20);
            this.tbTemplate.TabIndex = 1;
            // 
            // ReportCardWindowsGUIMerger
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(495, 225);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.grpMergeFrom);
            this.Controls.Add(this.btnMerge);
            this.Controls.Add(this.ssMergeStatus);
            this.Controls.Add(this.grpMergeType);
            this.Controls.Add(this.grpNames);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "ReportCardWindowsGUIMerger";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Report Cards";
            this.Load += new System.EventHandler(this.ReportCardWindowsGUIMerger_Load);
            this.grpNames.ResumeLayout(false);
            this.grpNames.PerformLayout();
            this.grpMergeType.ResumeLayout(false);
            this.grpMergeType.PerformLayout();
            this.ssMergeStatus.ResumeLayout(false);
            this.ssMergeStatus.PerformLayout();
            this.grpMergeFrom.ResumeLayout(false);
            this.grpMergeFrom.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbNames;
        private System.Windows.Forms.GroupBox grpNames;
        private System.Windows.Forms.GroupBox grpMergeType;
        private System.Windows.Forms.CheckBox cbMergeToPDF;
        private System.Windows.Forms.CheckBox cbMergeToPUB;
        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.StatusStrip ssMergeStatus;
        private System.Windows.Forms.ToolStripStatusLabel ssStatusText;
        private System.Windows.Forms.ToolStripProgressBar ssProgress;
        private System.Windows.Forms.OpenFileDialog fdTemplateOpen;
        private System.Windows.Forms.Label lblSaveTo;
        private System.Windows.Forms.FolderBrowserDialog fdSaveTo;
        private System.Windows.Forms.GroupBox grpMergeFrom;
        private System.Windows.Forms.Label lblDatasource;
        private System.Windows.Forms.Button btnTemplateBrowse;
        private System.Windows.Forms.Label lblTemplate;
        private System.Windows.Forms.Button btnSavetoBrowse;
        private System.Windows.Forms.Button btnDatasourceBrowse;
        private System.Windows.Forms.OpenFileDialog fdDatasourceOpen;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox cbSelectAll;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.TextBox tbSaveTo;
        private System.Windows.Forms.TextBox tbDatasource;
        private System.Windows.Forms.TextBox tbTemplate;
    }
}