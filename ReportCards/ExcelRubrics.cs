namespace SouthernCluster.ReportCards
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;

    public class ExcelWorksheet : ReportCardWorksheet
    {
        internal ExcelWorksheet()
            : base()
        {
        }

        internal ExcelWorksheet(string sheet, OleDbConnection conn)
        {
            this.BeginInit();

            OleDbDataAdapter conncmd;
            DataTable table = new DataTable();

            conncmd = new OleDbDataAdapter("SELECT * FROM [" + sheet + "]", conn);
            conncmd.Fill(table);
            conncmd.Dispose();

            this.Load(table);

            this.EndInit();
        }
    }

    public class ExcelWorkbook : ReportCardWorkbook
    {
        private List<string> GetSheets(OleDbConnection conn)
        {
            DataTable tables = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            List<string> sheets = new List<string>();
            for (int i = 0; i < tables.Rows.Count; i++)
            {
                sheets.Add(tables.Rows[i][2].ToString());
            }
            return sheets;
        }

        internal ExcelWorkbook(string name)
            : base(name)
        {
            string connstr;
            if (name.EndsWith(".xlsx"))
            {
                connstr = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                          "Data Source=" + name + ";" +
                          "Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=15;READONLY=TRUE\"";
            }
            else if (name.EndsWith(".xlsb"))
            {
                connstr = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                          "Data Source=" + name + ";" +
                          "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;MAXSCANROWS=15;READONLY=TRUE\"";
            }
            else if (name.EndsWith(".xls"))
            {
                connstr = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                          "Data Source=" + name + ";" +
                          "Extended Properties=\"Excel 8.0;HDR=YES;READONLY=TRUE\"";
            }
            else
            {
                throw new ArgumentException();
            }
            using (OleDbConnection conn = new OleDbConnection(connstr))
            {
                conn.Open();
                List<string> sheets = GetSheets(conn);
                if (sheets.Contains("Marks"))
                {
                    this.Add("Students", new ExcelWorksheet("Marks", conn));
                }
                else
                {
                    foreach (string sheet in sheets)
                    {
                        this.Add(sheet, new ExcelWorksheet(sheet, conn));
                    }
                }
            }
        }
    }

    public class ExcelMarksTable : ReportCardMarksTable
    {
        public ExcelMarksTable(string name)
            : base(name)
        {
        }

        protected override void Dispose(bool disposing)
        {
        }

        protected override ReportCardWorkbook GetWorkbook(string id)
        {
            return new ExcelWorkbook(id);
        }
    }
}
