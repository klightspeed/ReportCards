using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SouthernCluster.ReportCards
{
    public partial class DataTableGUI : Form
    {
        public Dictionary<string, bool> SelectedEntries { get; protected set; }

        public DataTableGUI(IEnumerable<DataRow> data, Dictionary<string, bool> selected)
        {
            InitializeComponent();
            SelectedEntries = selected;
            DataView dataView = data.OrderBy(row => row["Name"]).CopyToDataTable().AsDataView();
            srcDataTable.DataSource = dataView;
            gridDataTable.DataSource = srcDataTable;
            gridDataTable.DefaultCellStyle.NullValue = "----";
            gridDataTable.ReadOnly = true;
            /*
            gridDataTable.VirtualMode = true;
            DataGridViewCheckBoxColumn selectColumn = new DataGridViewCheckBoxColumn(false);
            selectColumn.TrueValue = true;
            selectColumn.FalseValue = false;
            gridDataTable.Columns.Insert(0, selectColumn);
            gridDataTable.Columns[0].ReadOnly = false;

            for (int i = 1; i < gridDataTable.Columns.Count; i++)
            {
                gridDataTable.Columns[i].ReadOnly = true;
            }

            foreach (DataGridViewRow row in gridDataTable.Rows)
            {
                row.Cells[0].Value = SelectedEntries[(string)row.Cells["Name"].Value];
            }
             */
        }

        private void DataTableGUI_Load(object sender, EventArgs e)
        {
            gridDataTable.AutoResizeColumns();
        }

        /*
        private void gridDataTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                string name = (string)gridDataTable.Rows[e.RowIndex].Cells["Name"].Value;
                object value = gridDataTable.Rows[e.RowIndex].Cells[0].Value;
                if (value != null)
                {
                    SelectedEntries[name] = (bool)value;
                }
            }
        }
         */
    }
}
