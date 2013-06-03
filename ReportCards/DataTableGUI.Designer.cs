namespace SouthernCluster.ReportCards
{
    partial class DataTableGUI
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
            this.components = new System.ComponentModel.Container();
            this.gridDataTable = new System.Windows.Forms.DataGridView();
            this.srcDataTable = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.gridDataTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.srcDataTable)).BeginInit();
            this.SuspendLayout();
            // 
            // gridDataTable
            // 
            this.gridDataTable.AllowUserToAddRows = false;
            this.gridDataTable.AllowUserToDeleteRows = false;
            this.gridDataTable.AllowUserToOrderColumns = true;
            this.gridDataTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridDataTable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridDataTable.Location = new System.Drawing.Point(0, 0);
            this.gridDataTable.Name = "gridDataTable";
            this.gridDataTable.Size = new System.Drawing.Size(784, 442);
            this.gridDataTable.TabIndex = 0;
            // 
            // DataTableGUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 442);
            this.Controls.Add(this.gridDataTable);
            this.Name = "DataTableGUI";
            this.Text = "DataTableGUI";
            this.Load += new System.EventHandler(this.DataTableGUI_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridDataTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.srcDataTable)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView gridDataTable;
        private System.Windows.Forms.BindingSource srcDataTable;
    }
}