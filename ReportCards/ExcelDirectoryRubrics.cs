using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Data.Odbc;

namespace SouthernCluster.ReportCards
{
    [System.ComponentModel.DesignerCategory("Code")]
    internal class ExcelDirectoryWorksheet : ReportCardWorksheet
    {
        private List<string> GetSheets(OdbcConnection conn)
        {
            DataTable tables = conn.GetSchema("Tables");
            List<string> sheets = new List<string>();
            for (int i = 0; i < tables.Rows.Count; i++)
            {
                sheets.Add(tables.Rows[i]["TABLE_NAME"].ToString());
            }
            return sheets;
        }

        private void LoadWorkbook(string filename)
        {
            string connstr = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + filename;

            using (OdbcConnection conn = new OdbcConnection(connstr))
            {
                conn.Open();
                List<string> sheets = GetSheets(conn);
                if (sheets.Contains("Marks"))
                {
                    LoadMarksSheet(filename, conn);
                }
            }
        }

        private void LoadMarksSheet(string filename, OdbcConnection conn)
        {
            OdbcDataAdapter conncmd;
            using (DataTable table = new DataTable())
            {
                using (OdbcCommandBuilder bldr = new OdbcCommandBuilder())
                {
                    using (conncmd = new OdbcDataAdapter("SELECT * FROM " + bldr.QuoteIdentifier("Marks", conn), conn))
                    {
                        conncmd.Fill(table);
                    }
                }

                if (table.Columns.Contains("Name"))
                {
                    LoadDataTable(filename, table);
                }
            }
        }

        private void LoadDataTable(string filename, DataTable table)
        {
            foreach (DataRow srcrow in table.Rows)
            {
                string rowname = srcrow["Name"].ToString();

                if (!String.IsNullOrEmpty(rowname))
                {
                    DataRow row;
                    if (this.Rows.Contains(rowname))
                    {
                        row = this.Rows.Find(rowname);
                    }
                    else
                    {
                        row = this.NewRow();
                        row["Name"] = rowname;
                        this.Rows.Add(row);
                    }

                    foreach (DataColumn column in table.Columns)
                    {
                        string colname = column.ColumnName;
                        string colvalue = (srcrow[column] ?? "").ToString();

                        if (!this.Columns.Contains(colname))
                        {
                            this.Columns.Add(colname);
                        }

                        if (row[colname] == null || row[colname] == DBNull.Value)
                        {
                            row[colname] = colvalue;
                        }
                        else if (row[colname].ToString().ToUpper() != colvalue.ToUpper())
                        {
                            throw new InvalidDataException(String.Format("Data file [{0}] student [{1}] column [{2}] has different value [{3}] to previous value [{4}]", filename, rowname, colname, colvalue, row[colname]));
                        }
                    }
                }
            }
        }
        
        public ExcelDirectoryWorksheet(string id)
        {
            this.BeginLoadData();
            
            this.Columns.Add("Name", typeof(string));
            this.PrimaryKey = new DataColumn[] { this.Columns["Name"] };

            List<string> columns = new List<string>();
            Dictionary<string, Dictionary<string, string>> data = new Dictionary<string, Dictionary<string, string>>();

            foreach (string filename in Directory.GetFiles(Path.GetDirectoryName(id), Path.GetFileName(id), SearchOption.TopDirectoryOnly))
            {
                if (filename.EndsWith(".xlsx") || filename.EndsWith(".xlsb") || filename.EndsWith(".xls") || filename.EndsWith(".xlsm"))
                {
                    LoadWorkbook(filename);
                }
            }

            this.EndLoadData();
        }
    }

    internal class ExcelDirectoryWorkbook : ReportCardWorkbook
    {
        public ExcelDirectoryWorkbook(string id)
            : base(id)
        {
            this.Add("Students", new ExcelDirectoryWorksheet(id));
        }
    }

    internal class ExcelDirectoryMarksTable : ReportCardMarksTable
    {
        public ExcelDirectoryMarksTable(string name)
            : base(name)
        {
        }

        protected override void Dispose(bool disposing)
        {
        }

        protected override ReportCardWorkbook GetWorkbook(string id)
        {
            return new ExcelDirectoryWorkbook(id);
        }
    }
}
