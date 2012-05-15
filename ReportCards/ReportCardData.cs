namespace SouthernCluster.ReportCards
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;

    internal class ReportCardWorksheet : DataTable
    {
        public ReportCardWorksheet()
            : base()
        {
        }

        protected void Load(string[][] data)
        {
            this.BeginLoadData();

            foreach (string col in data[0])
            {
                if (col != null && col != "")
                {
                    this.Columns.Add(col);
                }
            }

            this.PrimaryKey = new DataColumn[] { this.Columns[0] };

            for (int i = 1; i < data.Length; i++)
            {
                if (data[i][0] != null && data[i][0].Trim() != "")
                {
                    List<string> datarow = new List<string>();
                    for (int j = 0; j < data[0].Length; j++)
                    {
                        if (data[0][j] != null && data[0][j] != "")
                        {
                            datarow.Add(data[i][j]);
                        }
                    }
                    this.Rows.Add(datarow.ToArray());
                }
            }

            this.EndLoadData();
        }

        protected void Load(DataTable table)
        {
            string[][] data;
            int width = table.Columns.Count;
            int height = table.Rows.Count + 1;
            data = new string[height][];

            for (int row = 0; row < height; row++)
            {
                data[row] = new string[width];
            }

            for (int col = 0; col < width; col++)
            {
                data[0][col] = table.Columns[col].ColumnName;
            }

            for (int row = 0; row < height - 1; row++)
            {
                for (int col = 0; col < width; col++)
                {
                    data[row + 1][col] = table.Rows[row][col].ToString();
                }
            }
            this.Load(data);
        }

        public ReportCardWorksheet Transpose()
        {
            string[][] _data = new string[this.Columns.Count][];
            ReportCardWorksheet sheet = new ReportCardWorksheet();

            try
            {

                sheet.BeginInit();
                sheet.BeginLoadData();

                sheet.Columns.Add(this.Columns[0].ColumnName);
                for (int i = 0; i < this.Rows.Count; i++)
                {
                    sheet.Columns.Add(this.Rows[i][0] as string);
                }

                sheet.PrimaryKey = new DataColumn[] { sheet.Columns[0] };
                sheet.TableName = this.TableName;

                for (int i = 1; i < this.Columns.Count; i++)
                {
                    List<string> rowdata = new List<string>();
                    rowdata.Add(this.Columns[i].ColumnName);
                    for (int j = 0; j < this.Rows.Count; j++)
                    {
                        rowdata.Add(this.Rows[j][i] as string);
                    }
                    sheet.Rows.Add(rowdata.ToArray());
                }

                sheet.EndLoadData();
                sheet.EndInit();

                return sheet;
            }
            catch
            {
                sheet.Dispose();
                throw;
            }
        }
    }

    internal abstract class ReportCardWorkbook : Dictionary<string, ReportCardWorksheet>
    {
        public ReportCardWorkbook(string id)
        {
        }
    }

    internal abstract class ReportCardMarksTable : IDisposable
    {
        private string id;

        public ReportCardMarksTable(string id)
        {
            this.id = id;
        }

        ~ReportCardMarksTable()
        {
            Dispose(false);
        }

        public void Close()
        {
            Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected abstract void Dispose(bool disposing);

        protected abstract ReportCardWorkbook GetWorkbook(string id);

        public void Fill(DataTable table)
        {
            ReportCardWorkbook spreadsheet = GetWorkbook(id);

            table.BeginInit();
            table.BeginLoadData();
            table.Columns.Add("Name");
            table.PrimaryKey = new DataColumn[] { table.Columns["Name"] };

            List<DataColumn> columns = new List<DataColumn>();

            string[] sheetnames = new string[spreadsheet.Keys.Count];
            spreadsheet.Keys.CopyTo(sheetnames, 0);
            foreach (string title in sheetnames)
            {
                ReportCardWorksheet sheet = spreadsheet[title];
                if (sheet.Columns[0].ColumnName == "Abbreviation")
                {
                    spreadsheet[title] = sheet = sheet.Transpose();
                }

                foreach (DataColumn col in sheet.Columns)
                {
                    string colname = col.ColumnName;

                    if (title != "Students")
                    {
                        if (sheet.Columns[0].ColumnName == "Name" && sheet.Columns.Count == 2 && colname != "Name")
                        {
                            colname = title.Replace(' ', '_');
                        }
                        else
                        {
                            colname = title.Replace(' ', '_') + "_" + colname;
                        }
                    }

                    if (col.ColumnName != "Abbreviation" && col.ColumnName != "Name" && !col.ColumnName.StartsWith("#") && !table.Columns.Contains(colname))
                    {
                        DataColumn tcol = new DataColumn(colname);
                        if (sheet.Columns[0].ColumnName == "Abbreviation" &&
                            sheet.Rows[0][0] as string == "Indicator")
                        {
                            tcol.Caption = sheet.Rows[0][col] as string;
                        }

                        table.Columns.Add(tcol);
                        columns.Add(col);
                    }
                }
            }

            foreach (DataRow row in spreadsheet["Students"].Rows)
            {
                string[] rowdata = new string[columns.Count + 1];
                rowdata[0] = row["Name"] as string;
                for (int i = 0; i < columns.Count; i++)
                {
                    DataColumn col = columns[i];
                    DataRow trow = col.Table.Rows.Find(row["Name"]);
                    if (trow != null)
                    {
                        rowdata[i+1] = trow[col] as string;
                    }
                    else
                    {
                        rowdata[i+1] = null;
                    }
                }

                table.Rows.Add(rowdata);
            }

            table.EndLoadData();
            table.EndInit();
        }
    }

    internal class ReportCardData : IDisposable, IEnumerable<DataRow>
    {
        private DataTable table;

        public ReportCardData(string name)
        {
            ReportCardMarksTable marks = null;

            try
            {
                table = new DataTable();

                if (name.StartsWith("grubrics:"))
                {
                    marks = new GoogleMarksTable(name.Substring(9));
                }
                else if (name.EndsWith(".xlsx") || name.EndsWith(".xlsb") || name.EndsWith(".xls"))
                {
                    marks = new ExcelMarksTable(name);
                }
                else
                {
                    throw new ArgumentException();
                }

                marks.Fill(table);
            }
            finally
            {
                if (marks != null)
                {
                    marks.Dispose();
                }
            }
        }

        ~ReportCardData()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected void Dispose(bool disposing)
        {
            if (table != null)
            {
                table.Dispose();
            }

            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

        public DataRow this[string name]
        {
            get
            {
                string filter = "Name = '" + name.Replace("'", "''") + "'";
                using (DataView view = new DataView(table, filter, "Name", DataViewRowState.CurrentRows))
                {
                    DataRow row = view[0].Row;
                    return row;
                }
            }
        }

        public DataRow this[int index]
        {
            get
            {
                return table.Rows[index];
            }
        }

        public int Count
        {
            get
            {
                return table.Rows.Count;
            }
        }

        public string[] Names
        {
            get
            {
                List<string> ret = new List<string>();
                foreach (DataRow row in table.Rows)
                {
                    ret.Add(row["Name"].ToString());
                }
                ret.Sort();
                return ret.ToArray();
            }
        }


        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<DataRow> GetEnumerator()
        {
            foreach (DataRow row in table.Rows)
            {
                yield return row;
            }
        }
    }
}
