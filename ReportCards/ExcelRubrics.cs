namespace SouthernCluster.ReportCards
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.Odbc;
    using System.IO;

    internal class ExcelWorksheet : ReportCardWorksheet
    {
        public ExcelWorksheet()
            : base()
        {
        }

        public ExcelWorksheet(string sheet, OdbcConnection conn)
        {
            //this.BeginInit();

            OdbcDataAdapter conncmd;
            using (DataTable table = new DataTable())
            {
                using (OdbcCommandBuilder bldr = new OdbcCommandBuilder())
                {
                    using (conncmd = new OdbcDataAdapter("SELECT * FROM " + bldr.QuoteIdentifier(sheet, conn), conn))
                    {
                        conncmd.Fill(table);
                    }
                }

                this.Load(table);
            }

            //this.EndInit();
        }
    }

    internal class ExcelWorkbook : ReportCardWorkbook
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

        public ExcelWorkbook(string name)
            : base(name)
        {
            string connstr;
            if (name.EndsWith(".xlsx") || name.EndsWith(".xlsb") || name.EndsWith(".xls") || name.EndsWith(".xlsm"))
            {
                connstr = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + name;
            }
            else
            {
                throw new ArgumentException();
            }
            using (OdbcConnection conn = new OdbcConnection(connstr))
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

    internal class ExcelMarksTable : ReportCardMarksTable
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
