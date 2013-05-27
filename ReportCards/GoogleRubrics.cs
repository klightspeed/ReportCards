namespace SouthernCluster.ReportCards
{
    using System;
    using System.Net;
    using System.Text;
    using System.Xml;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Threading;

    [System.ComponentModel.DesignerCategory("Code")]
    internal class GoogleWorksheet : ReportCardWorksheet
    {
        public GoogleWorksheet()
            : base()
        {
        }

        public GoogleWorksheet(WebClient client, string url)
            : base()
        {
            //this.BeginInit();

            XmlDocument rootnode = new XmlDocument();
            rootnode.LoadXml(client.DownloadString(url));
            this.TableName = rootnode.GetElementsByTagName("title", "http://www.w3.org/2005/Atom")[0].ChildNodes[0].Value;
            int height = Int32.Parse(rootnode.GetElementsByTagName("rowCount", "http://schemas.google.com/spreadsheets/2006")[0].ChildNodes[0].Value);
            int width = Int32.Parse(rootnode.GetElementsByTagName("colCount", "http://schemas.google.com/spreadsheets/2006")[0].ChildNodes[0].Value);

            string[][] _data = new string[height][];
            
            for (int row = 0; row < height; row++)
            {
                _data[row] = new string[width];
            }

            foreach (XmlNode n_cell in rootnode.GetElementsByTagName("cell", "http://schemas.google.com/spreadsheets/2006"))
            {
                int row = Int32.Parse(n_cell.Attributes["row"].Value);
                int col = Int32.Parse(n_cell.Attributes["col"].Value);
                string val = null;
                if (n_cell.HasChildNodes)
                {
                    val = n_cell.ChildNodes[0].Value;
                }
                _data[row - 1][col - 1] = val;
            }

            this.Load(_data);

            //this.EndInit();
        }
    }

    internal class GoogleWorkBook : ReportCardWorkbook
    {
        public GoogleWorkBook(string spreadsheet)
            : base(spreadsheet)
        {
            using (WebClient client = new WebClient())
            {
                string feed = "https://spreadsheets.google.com/feeds/worksheets/" + spreadsheet + "/public/values";
                XmlDocument rootnode = new XmlDocument();
                rootnode.LoadXml(client.DownloadString(feed));
                foreach (XmlNode n_entry in rootnode.GetElementsByTagName("entry", "http://www.w3.org/2005/Atom"))
                {
                    XmlElement e_entry = (XmlElement)n_entry;
                    string title = e_entry.GetElementsByTagName("title", "http://www.w3.org/2005/Atom")[0].ChildNodes[0].Value;
                    foreach (XmlNode n_link in e_entry.GetElementsByTagName("link", "http://www.w3.org/2005/Atom"))
                    {
                        if (n_link.Attributes["rel"].Value == "http://schemas.google.com/spreadsheets/2006#cellsfeed")
                        {
                            string href = n_link.Attributes["href"].Value;
                            Console.WriteLine("Getting spreadsheet {0} = {1}", title, href);
                            this.Add(title, new GoogleWorksheet(client, href));
                        }
                    }
                }
            }
        }
    }

    internal class GoogleMarksTable : ReportCardMarksTable
    {
        public GoogleMarksTable(string id) : base(id) { }

        protected override void Dispose(bool disposing)
        {
        }

        protected override ReportCardWorkbook GetWorkbook(string id)
        {
            return new GoogleWorkBook(id);
        }
    }
}
