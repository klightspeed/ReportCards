namespace SouthernCluster.ReportCards
{
    using System;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Xps;
    using System.Windows.Xps.Packaging;
    using Microsoft.Office.Interop.Publisher;
    using Publisher = Microsoft.Office.Interop.Publisher;
    using Microsoft.Office.Core;

    internal class ReportCardRangeCollection : List<ReportCardRange>
    {
        private Dictionary<string, string> variables;

        public ReportCardRangeCollection()
            : base()
        {
            this.variables = new Dictionary<string, string>();
        }

        public string GetVar(string key)
        {
            if (variables.ContainsKey(key.ToUpper()))
            {
                return variables[key.ToUpper()];
            }
            else
            {
                return null;
            }
        }

        public bool HasVar(string key)
        {
            return variables.ContainsKey(key.ToUpper());
        }

        public void SetVar(string key, string val)
        {
            variables[key.ToUpper()] = val.Replace("`n", "\n");
        }

        public IEnumerable<KeyValuePair<string, string>> Vars
        {
            get
            {
                foreach (KeyValuePair<string, string> var in variables)
                {
                    yield return var;
                }
            }
        }
    }

    public class InvalidReportCardRangeException : InvalidOperationException
    {
        public ReportCardRange Range;

        public InvalidReportCardRangeException(ReportCardRange range, string message)
            : base(message + "\nSpecifier string was:\n" + range.specifierstring)
        {
            this.Range = range;
        }
    }

    public class ReportCardRange
    {
        public int col = 0;
        public int row = 0;
        public int width = 0;
        public int height = 0;
        public int tblno = 0;
        public string tablename = null;
        public string marktype = null;
        public string[] marks = null;
        public string fieldname = null;
        public string[] colours = { "white", "black", "black", "white" };
        public string specifierstring = null;

        internal static ReportCardRangeCollection ParseCommands(Publisher.Document doc)
        {
            ReportCardRangeCollection ret = null;
            foreach (Publisher.Shape shape in doc.ScratchArea.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    string commandstring = shape.TextFrame.TextRange.Story.TextRange.Text;
                    ret = ParseCommands(commandstring);
                }
                if (ret != null)
                {
                    return ret;
                }
            }
            return null;
        }

        internal static ReportCardRangeCollection ParseCommands(string commandstring)
        {
            ReportCardRangeCollection ranges = new ReportCardRangeCollection();
            foreach (string command in commandstring.Split(new char[] { '\n', '\r' }))
            {
                if (!command.StartsWith("#"))
                {
                    int colonpos = command.IndexOf(":");
                    int equalpos = command.IndexOf("=");
                    if (command.Contains("=") && (colonpos == -1 || equalpos < colonpos))
                    {
                        Console.WriteLine("Parsing variable: {0}\n", command);
                        KeyValuePair<string, string> var = ReportCardRange.ParseVariable(command);
                        if (var.Key != null)
                        {
                            ranges.SetVar(var.Key, var.Value);
                        }
                    }
                    else if (command.Contains(":"))
                    {
                        Console.WriteLine("Parsing command: {0}", command);
                        ReportCardRange range = ReportCardRange.ParseCommand(command, ranges);
                        if (range != null)
                        {
                            ranges.Add(range);
                        }
                    }
                }
            }
            if (ranges.Count >= 1 || ranges.Vars.Count() >= 1)
            {
                return ranges;
            }
            else
            {
                return null;
            }
        }

        internal static KeyValuePair<string, string> ParseVariable(string commandstring)
        {
            string[] args = commandstring.Split(new char[] { '=' }, 2);
            return new KeyValuePair<string, string>(args[0].ToUpper(), args[1]);
        }

        internal static ReportCardRange ParseCommand(string commandstring, ReportCardRangeCollection ranges)
        {
            return new ReportCardRange(commandstring, ranges);
        }

        private ReportCardRange(string commandstring, ReportCardRangeCollection ranges)
        {
            specifierstring = commandstring;
            string[] args = commandstring.Split(new char[] { ':' });
            marktype = args[0].ToUpper().Trim();
            Int32.TryParse(args[1], out tblno);
            tablename = args[1];
            Int32.TryParse(args[2], out col);
            Int32.TryParse(args[3], out row);
            Int32.TryParse(args[4], out width);
            Int32.TryParse(args[5], out height);
            marks = args[6].Trim().Split(new char[] { ',' });
            fieldname = args[7].Trim();

            if (ranges.HasVar("unshaded-colour"))
            {
                string[] unshaded = ranges.GetVar("unshaded-colour").Split(new char[] { ',' });
                colours[0] = unshaded[1].Trim();
                colours[1] = unshaded[0].Trim();
            }

            if (ranges.HasVar("shaded-colour"))
            {
                string[] shaded = ranges.GetVar("shaded-colour").Split(new char[] { ',' });
                colours[2] = shaded[1].Trim();
                colours[3] = shaded[0].Trim();
            }

            if (col <= 0)
            {
                throw new InvalidReportCardRangeException(this, "Column is less than or equal to zero");
            }
            else if (row <= 0)
            {
                throw new InvalidReportCardRangeException(this, "Row is less than or equal to zero");
            }
            else if (marktype != "TICK" && marktype != "SHADE" && marktype != "TABLENAME")
            {
                throw new InvalidReportCardRangeException(this, String.Format("Unrecognized range type {0}", marktype));
            }
            else if (marktype != "TABLENAME" && fieldname == "")
            {
                throw new InvalidReportCardRangeException(this, "Field name is empty");
            }
            else if (marktype != "TABLENAME" && width != marks.Length && height != marks.Length)
            {
                throw new InvalidReportCardRangeException(this, "Size of selection does not match number of marks");
            }
            else if (marktype != "TABLENAME" && width != 1 && height != 1)
            {
                throw new InvalidReportCardRangeException(this, "Selection must be a single row or a single column");
            }
        }
    }

    internal interface IReportCardMergeEntry
    {
        void Update(DataRow data);
        void Reset();
        string Name { get; }
        string Value { get; set; }
    }

    internal class ReportCardMergeField : IReportCardMergeEntry
    {
        private Publisher.TextRange range;
        private string fieldname;
        public Publisher.Font font;

        public ReportCardMergeField(Publisher.TextRange range, string fieldname)
        {
            this.range = range;
            this.fieldname = fieldname;
            this.font = range.MajorityFont;
        }

        public void Update(DataRow data)
        {
            string val = null;
            if (data.Table.Columns.Contains(fieldname))
            {
                val = data[fieldname].ToString();
            }
            Value = val;
        }

        public void Reset()
        {
            Value = "«" + fieldname + "»";
        }

        public string Name
        {
            get
            {
                return fieldname;
            }
        }

        public string Value
        {
            get
            {
                return range.Text;
            }
            set
            {
                Publisher.TextRange delrange = range.Duplicate;
                range.Collapse(Publisher.PbCollapseDirection.pbCollapseStart);
                delrange.Delete();
                if (value != null)
                {
                    string[] paras = value.Split('\n');
                    range.InsertAfter(paras[0].Trim());
                    for (int i = 1; i < paras.Length; i++)
                    {
                        range.InsertAfter("\r");
                        range.InsertAfter(paras[i].Trim());
                    }
                    range.Font = font;
                }
            }
        }
    }

    internal class ReportCardTickRange: IReportCardMergeEntry
    {
        private Dictionary<string,Publisher.Cell> cells;
        private string starcell;
        private Publisher.Font starcellfont;
        private string fieldname;
        private string mark;
        private bool usewingdingticks;

        public ReportCardTickRange (
            Dictionary<string,Publisher.Cell> cells,
            string fieldname,
            bool usewingdingticks
        ){
            this.cells = new Dictionary<string,Publisher.Cell>();
            this.fieldname = fieldname;
            this.usewingdingticks = usewingdingticks;
            this.mark = "";
            foreach (KeyValuePair<string, Publisher.Cell> kvp in cells)
            {
                string key = kvp.Key;
                
                if (key.EndsWith("*"))
                {
                    key = key.Replace("*", "");
                    starcell = key;
                    starcellfont = kvp.Value.TextRange.MajorityFont;
                }

                this.cells[key] = kvp.Value;
            }
        }

        public void Update (DataRow data)
        {
            string val = null;
            if (data.Table.Columns.Contains(fieldname))
            {
                val = data[fieldname].ToString().ToUpper();
            }
            Value = val;
        }

        public void Reset()
        {
            Value = "";
        }

        public string Value
        {
            get
            {
                return mark;
            }
            set
            {
                if (mark != "")
                {
                    if (cells.ContainsKey(mark) && cells[mark] != null)
                    {
                        cells[mark].TextRange.Delete();
                    }
                    else if (starcell != null && cells.ContainsKey(starcell) && cells[starcell] != null)
                    {
                        cells[starcell].TextRange.Delete();
                    }
                }

                if (value != null)
                {
                    mark = value.Trim().ToUpper();
                    if (cells.ContainsKey(mark) && cells[mark] != null)
                    {
                        cells[mark].TextRange.Delete();
                        if (usewingdingticks)
                        {
                            cells[mark].TextRange.Text = "ü"; // Wingdings Check Mark
                            cells[mark].TextRange.Font.SetScriptName(Publisher.PbFontScriptType.pbFontScriptAsciiLatin, "Wingdings");
                        }
                        else
                        {
                            cells[mark].TextRange.Text = "✓"; // Unicode Check Mark
                        }
                    }
                    else if (starcell != null && cells.ContainsKey(starcell) && cells[starcell] != null)
                    {
                        cells[starcell].TextRange.Delete();
                        cells[starcell].TextRange.Text = mark;
                        cells[starcell].TextRange.Font = starcellfont;
                    }
                    else
                    {
                        mark = "";
                    }
                }
                else
                {
                    mark = "";
                }
            }
        }

        public string Name
        {
            get
            {
                return fieldname;
            }
        }
    }

    internal class ReportCardShadeRange : IReportCardMergeEntry
    {
        private Dictionary<string, Publisher.Cell> cells;
        private string fieldname;
        private string mark;
        private int unshaded_bgcolour;
        private int unshaded_fgcolour;
        private int shaded_bgcolour;
        private int shaded_fgcolour;

        private int ParseColour(string name, int defcolour)
        {
            try
            {
                return System.Drawing.ColorTranslator.FromHtml(name).ToArgb() & 0x00FFFFFF;
            }
            catch
            {
                return defcolour;
            }
        }

        public ReportCardShadeRange(
            Dictionary<string, Publisher.Cell> cells,
            string fieldname,
            string[] colours
        )
        {
            this.cells = cells;
            this.fieldname = fieldname;
            this.mark = "";

            unshaded_bgcolour = ParseColour(colours[0], 0xFFFFFF);
            unshaded_fgcolour = ParseColour(colours[1], 0x000000);
            shaded_bgcolour = ParseColour(colours[2], 0x000000);
            shaded_fgcolour = ParseColour(colours[3], 0xFFFFFF);
            foreach (Publisher.Cell cell in this.cells.Values)
            {
                cell.Fill.BackColor.RGB = unshaded_bgcolour;
                cell.Fill.BackColor.Transparency = 1;
                cell.Fill.ForeColor.RGB = unshaded_bgcolour;
                cell.Fill.ForeColor.Transparency = 1;
                cell.TextRange.Font.Color.RGB = unshaded_fgcolour;
            }
        }

        public void Update(DataRow data)
        {
            string val = null;
            if (data.Table.Columns.Contains(fieldname))
            {
                val = data[fieldname].ToString().ToUpper();
            }
            Value = val;
        }

        public void Reset()
        {
            Value = "";
        }

        public string Value
        {
            get
            {
                return mark;
            }
            set
            {
                if (mark != "" && cells.ContainsKey(mark) && cells[mark] != null)
                {
                    cells[mark].Fill.BackColor.RGB = unshaded_bgcolour;
                    cells[mark].Fill.BackColor.Transparency = 1;
                    cells[mark].Fill.ForeColor.RGB = unshaded_bgcolour;
                    cells[mark].Fill.ForeColor.Transparency = 1;
                    cells[mark].TextRange.Font.Color.RGB = unshaded_fgcolour;
                }

                if (value != null)
                {
                    mark = value.Trim().ToUpper();
                    if (cells.ContainsKey(mark) && cells[mark] != null)
                    {
                        cells[mark].Fill.BackColor.RGB = shaded_bgcolour;
                        cells[mark].Fill.BackColor.Transparency = 0;
                        cells[mark].Fill.ForeColor.RGB = shaded_bgcolour;
                        cells[mark].Fill.ForeColor.Transparency = 0;
                        cells[mark].TextRange.Font.Color.RGB = shaded_fgcolour;
                    }
                    else
                    {
                        mark = "";
                    }
                }
                else
                {
                    mark = "";
                }
            }
        }

        public string Name
        {
            get
            {
                return fieldname;
            }
        }
    }

    internal class ReportCardPicture : IReportCardMergeEntry
    {
        private string fieldname;
        private string filename;
        private float width;
        private float height;
        private float x;
        private float y;
        private float contrast;
        private float brightness;
        private Object parent;
        private Publisher.Shape shape;

        public ReportCardPicture (Publisher.Shape shape)
        {
            this.shape = shape;
            this.parent = shape.Parent;
            this.fieldname = shape.AlternativeText.Trim();
            if (this.fieldname != null && 
                this.fieldname.StartsWith("«") &&
                this.fieldname.Contains("»"))
            {
                this.fieldname = this.fieldname.Substring(1, this.fieldname.IndexOf("»") - 1);
                this.width = (float)shape.Width;
                this.height = (float)shape.Height;
                this.x = (float)shape.Top;
                this.y = (float)shape.Left;
                this.contrast = (float)shape.PictureFormat.Contrast;
                this.brightness = (float)shape.PictureFormat.Brightness;
                if (shape.PictureFormat.IsEmpty == MsoTriState.msoFalse)
                {
                    this.filename = shape.PictureFormat.Filename;
                }
                else
                {
                    this.filename = null;
                }
            }
        }

        public void Update(DataRow data)
        {
            string val = null;
            if (fieldname != null)
            {
                if (data.Table.Columns.Contains(fieldname))
                {
                    val = data[fieldname].ToString();
                }
                Value = val;
            }
        }

        public void Reset()
        {
            Value = null;
        }

        public string Name 
        {
            get 
            {
                return fieldname;
            }
        }

        public string Value 
        { 
            get 
            {
                return filename;
            }

            set 
            {
                if (value != null && File.Exists(value))
                {
                    //string tempname = Path.GetTempFileName();
                    //string temppicname = tempname + value.Substring(value.LastIndexOf("."));
                    //File.Copy(value, temppicname, true);
                    shape.PictureFormat.RestoreOriginalColors();
                    shape.PictureFormat.Replace(value, Publisher.PbPictureInsertAs.pbPictureInsertAsLinked);
                    filename = value;
                }
                else
                {
                    shape.Fill.BackColor.RGB = 0x00FFFFFF;
                    shape.PictureFormat.Recolor(shape.Fill.BackColor, MsoTriState.msoFalse);
                    filename = null;
                }
            }
        }
    }

    internal class ReportCardTable
    {
        int width;
        int height;
        Cell[] cells;
        int page;
        float x;
        float y;
        float tx;
        float ty;

        public ReportCardTable(Table table)
        {
            width = table.Cells.Width;
            height = table.Cells.Height;
            cells = new Cell[width * height];
            foreach (Cell cell in table.Cells){
                this[cell.Row, cell.Column] = cell;
            }
            Object o = table;
            List<Object> objs = new List<Object>();
            objs.Add(o);
            Publisher.Shape tblshp = null;
            Publisher.Shape shp = null;
            Publisher.Page pg = null;
            while ((o as Publisher.Application) == null)
            {
                o = o.GetType().InvokeMember("Parent", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty, null, o, null);
                objs.Add(o);
                shp = o as Publisher.Shape ?? shp;
                if (tblshp == null){
                    tblshp = shp;
                }
                pg = o as Publisher.Page ?? pg;
            }
            page = pg.PageIndex;

            x = (float)shp.Left * 2.54F / 72F;
            y = (float)shp.Top * 2.54F / 72F;
            tx = (float)tblshp.Left * 2.54F / 72F;
            ty = (float)tblshp.Top * 2.54F / 72F;
            Console.WriteLine("Table {0} [{1}] is at ({2},{3}) [({4},{5})] on page {6}, and is {7}x{8} cells",
                              tblshp.ID, shp.ID, tx, ty, x, y, page, width, height);
        }

        public Cell this[int row, int col]
        {
            get 
            {
                return cells[(row - 1) * width + col - 1];
            }
            private set
            {
                cells[(row - 1) * width + col - 1] = value;
            }
        }

        public int Width 
        {
            get 
            {
                return width;
            }
        }

        public int Height 
        {
            get 
            {
                return height;
            }
        }

        public int Page
        {
            get
            {
                return page;
            }
        }

        public float X
        {
            get
            {
                return x;
            }
        }

        public float Y
        {
            get
            {
                return y;
            }
        }
    }
    
    internal class ReportCard : IDisposable
    {
        private Page[] pages;
        private Page[] masterpages;
        private List<IReportCardMergeEntry> mergeentries;
        private Dictionary<int,ReportCardTable> cells;
        private Dictionary<string, int> tablenames;
        private bool deleted;
        private ReportCardRangeCollection ranges;
        private Publisher.Application pubapp;
        private Document doc;
        private ReportCardData data;
        private string curname;
        private string datasourcename;
        private string datafilter;

        public ReportCard(Publisher.Application pubapp, Document doc, string pubname, bool usewingdingticks)
        {
            this.pubapp = pubapp;
            this.doc = doc;
            this.mergeentries = null;
            this.pages = new Page[doc.Pages.Count];
            this.masterpages = new Page[doc.MasterPages.Count];
            this.ranges = ReportCardRange.ParseCommands(doc);
            this.datasourcename = this.ranges.GetVar("DATASOURCE");
            this.datafilter = this.ranges.GetVar("DATAFILTER");

            if (this.datasourcename == null)
            {
                string basename;

                if (pubname.EndsWith(".pub"))
                {
                    basename = pubname.Substring(0, pubname.Length - 4);
                }
                else
                {
                    basename = pubname;
                }

                if (File.Exists(basename + ".xlsx"))
                {
                    this.datasourcename = basename + ".xlsx";
                }
                else if (File.Exists(basename + ".xls"))
                {
                    this.datasourcename = basename + ".xls";
                }
                else if (File.Exists(basename + ".xlsb"))
                {
                    this.datasourcename = basename + ".xlsb";
                }
            }

            if (this.datasourcename != null && !(this.datasourcename.Contains("\\") || this.datasourcename.Contains("/") || this.datasourcename.Contains(":")))
            {
                this.datasourcename = Path.Combine(Path.GetDirectoryName(pubname), this.datasourcename);
            }

            int i = 0;

            foreach (Page page in doc.MasterPages)
            {
                masterpages[i++] = page;
            }

            i = 0;
            foreach (Page page in doc.Pages)
            {
                pages[i++] = page;
            }
            deleted = false;
            GetMergeFields();
            GetMarkRanges(usewingdingticks);
        }

        ~ReportCard()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (pubapp != null)
                {
                    if (doc != null)
                    {
                        doc.Close();
                        doc = null;
                    }

                    ((_Application)pubapp).Quit();
                }
                pubapp = null;
                doc = null;
            }

            data = null;
            deleted = true;
        }

        private void GetTableCellsInShape(Publisher.Shape shape){
            int id = shape.ID;
            Table table = shape.Table;

            tablenames.Add(shape.Name, shape.ID);
            cells.Add(id, new ReportCardTable(table));
        }

        private void GetMarkRanges(bool usewingdingticks){
            int curtblno = 0;
            foreach (ReportCardRange range in ranges)
            {
                if (range.tblno == 0 && range.tablename != "" && tablenames.ContainsKey(range.tablename))
                {
                    range.tblno = tablenames[range.tablename];
                }
                else if (curtblno != 0)
                {
                    range.tblno = curtblno;
                }

                if (range.marktype == "TABLENAME")
                {
                    Console.WriteLine("Trying to find table for tablename\n");
                    curtblno = 0;
                    foreach (KeyValuePair<int, ReportCardTable> table in cells)
                    {
                        if (range.row <= table.Value.Height &&
                            (range.height == table.Value.Height || range.height == 0) &&
                            range.col <= table.Value.Width &&
                            (range.width == table.Value.Width || range.width == 0) &&
                            table.Value[range.row, range.col] != null &&
                            table.Value[range.row, range.col].TextRange.Text.ToUpper().Contains(range.fieldname.ToUpper()))
                        {
                            bool matches = true;
                            foreach (string mark in range.marks)
                            {
                                string[] opt = mark.Split(new char[] { '=' }, 2);
                                if (opt[0].ToLower() == "p")
                                {
                                    int page;
                                    Int32.TryParse(opt[1], out page);
                                    if (table.Value.Page != page)
                                    {
                                        matches = false;
                                        break;
                                    }
                                }
                                else if (opt[0].ToLower() == "x")
                                {
                                    float x;
                                    Single.TryParse(opt[1], out x);
                                    if (table.Value.X < x - 1 || table.Value.X > x + 1)
                                    {
                                        matches = false;
                                        break;
                                    }
                                }
                               else if (opt[0].ToLower() == "y")
                                {
                                    float y;
                                    Single.TryParse(opt[1], out y);
                                    if (table.Value.Y < y - 1 || table.Value.Y > y + 1)
                                    {
                                        matches = false;
                                        break;
                                    }
                                }
                            }
                            if (matches)
                            {
                                if (range.tablename != "")
                                {
                                    tablenames.Add(range.tablename, table.Key);
                                }
                                curtblno = table.Key;
                                break;
                            }
                        }
                    }
                }
                else if (cells.ContainsKey(range.tblno) &&
                    range.row + range.height <= cells[range.tblno].Height + 1 &&
                    range.col + range.width <= cells[range.tblno].Width + 1)
                {
                    Dictionary<string, Cell> markcells = new Dictionary<string, Cell>();
                    int row = range.row;
                    int col = range.col;
                    for (int i = 0; i < range.marks.Length; i++)
                    {
                        Cell cell = cells[range.tblno][row, col];
                        if (cell == null)
                        {
                            throw new InvalidReportCardRangeException(range, String.Format(
                                "Unable to retrieve cell in row {0} column {1} of table {2}",
                                row, col, range.tblno
                            ));
                        }
                        markcells.Add(range.marks[i], cell);
                        if (range.width == 0)
                        {
                            row++;
                        }
                        else
                        {
                            col++;
                        }
                    }
                    if (range.marktype == "TICK")
                    {
                        mergeentries.Add(new ReportCardTickRange(markcells, range.fieldname, usewingdingticks));
                    }
                    else if (range.marktype == "SHADE")
                    {
                        mergeentries.Add(new ReportCardShadeRange(markcells, range.fieldname, range.colours));
                    }
                }
            }
        }
        
        private void GetMergeFields()
        {
            tablenames = new Dictionary<string, int>();
            cells = new Dictionary<int, ReportCardTable>();
            mergeentries = new List<IReportCardMergeEntry>();
            foreach (Publisher.Page page in pages)
            {
                foreach (Publisher.Shape shape in page.Shapes)
                {
                    GetMergeFieldsInShape(shape);
                }
            }
            if (masterpages != null)
            {
                foreach (Publisher.Page page in masterpages)
                {
                    foreach (Publisher.Shape shape in page.Shapes)
                    {
                        GetMergeFieldsInShape(shape);
                    }
                }
            }
        }

        private void GetMergeFieldsInShape(Publisher.Shape shape)
        {
            switch (shape.Type){
                case PbShapeType.pbTable:
                    GetTableCellsInShape(shape);
                    foreach (Cell cell in shape.Table.Cells)
                    {
                        if (cell.HasText)
                        {
                            GetMergeFieldsInTextRange(cell.TextRange);
                        }
                    }
                    break;
                case PbShapeType.pbTextFrame:
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        GetMergeFieldsInTextRange(shape.TextFrame.TextRange);
                    }
                    break;
                case PbShapeType.pbGroup:
                    foreach (Publisher.Shape grpshape in shape.GroupItems)
                    {
                        GetMergeFieldsInShape(grpshape);
                    }
                    break;
                case PbShapeType.pbBarCodePictureHolder:
                    goto case PbShapeType.pbPicture;
                case PbShapeType.pbLinkedPicture:
                    goto case PbShapeType.pbPicture;
                case PbShapeType.pbPlaceholder:
                    goto case PbShapeType.pbPicture;
                case PbShapeType.pbPicture:
                    // Cannot get merge field name; try Alternative text
                    Publisher.PictureFormat pic = shape.PictureFormat;
                    if (shape.AlternativeText.StartsWith("«") &&
                        shape.AlternativeText.Contains("»"))
                    {
                        mergeentries.Add(new ReportCardPicture(shape));
                    }
                    break;
            }
        }

        private void GetMergeFieldsInTextRange(TextRange range)
        {
            int fldstart = 0;
            int namestart;
            int nameend;
            int fldend;
            TextRange rng = range.Duplicate;

            while (rng.Text.Length >= 3)
            {
                string text = range.Text;
                int i = 0;
                int j = 0;

                i = text.IndexOf("«", fldstart);
                j = text.IndexOf("<<", fldstart);
                if (i >= 0 && (i < j || j < 0))
                {
                    fldstart = i;
                    namestart = i + 1;
                }
                else if (j >= 0)
                {
                    fldstart = j;
                    namestart = j + 2;
                }
                else
                {
                    break;
                }
                i = text.IndexOf("»", fldstart);
                j = text.IndexOf(">>", fldstart);
                if (i >= 0 && (i < j || j < 0))
                {
                    fldend = i + 1;
                    nameend = i;
                }
                else if (j >= 0)
                {
                    fldend = j + 2;
                    nameend = j;
                }
                else
                {
                    break;
                }
                TextRange txt = range.Characters(fldstart + 1, fldend - fldstart);
                string fieldname = text.Substring(namestart, nameend - namestart);
                mergeentries.Add(new ReportCardMergeField(txt, fieldname));
                fldstart = fldend;
            }

            foreach (Publisher.Shape shape in range.InlineShapes)
            {
                GetMergeFieldsInShape(shape);
            }
        }

        public void Update(DataRow row)
        {
            if (deleted)
            {
                throw new ObjectDisposedException("ReportCard");
            }

            doc.Application.ScreenUpdating = false;
            foreach (IReportCardMergeEntry entry in mergeentries)
            {
                if (ranges.HasVar(entry.Name))
                {
                    entry.Value = ranges.GetVar(entry.Name);
                }
                else
                {
                    entry.Update(row);
                }
            }
            doc.Application.ScreenUpdating = true;

            this.curname = row["Name"].ToString();
        }

        public void Reset()
        {
            if (deleted)
            {
                throw new ObjectDisposedException("ReportCard");
            }
            foreach (IReportCardMergeEntry entry in mergeentries)
            {
                entry.Reset();
            }
        }

        public Page[] Pages
        {
            get
            {
                if (deleted)
                {
                    throw new ObjectDisposedException("ReportCard");
                }
                return pages;
            }
        }

        private string FileNameEscape(string name)
        {
            name.Replace("%", "%25");
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c.ToString(), "%" + ((int)c).ToString("X2"));
            }
            return name;
        }

        public void MergeReport(string name)
        {
            if (deleted)
            {
                throw new ObjectDisposedException("ReportCards");
            }

            Update(data[name]);
            doc.Save();
        }

        public void SavePDF(string savepath){
            string tempoutpdfname = Path.GetTempFileName();
            string outbasename = FileNameEscape(curname);
            string outpdfname = Path.GetFullPath(savepath) + Path.DirectorySeparatorChar + outbasename + ".pdf";

            try
            {
                doc.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypePDF, tempoutpdfname, PbFixedFormatIntent.pbIntentStandard);
                File.Copy(tempoutpdfname, outpdfname, true);
            }
            finally
            {
                File.Delete(tempoutpdfname);
            }
        }

        public XpsDocument ExportXPS()
        {
            string tempoutxpsname = Path.GetTempFileName();
            string tempcopyxpsname = Path.GetTempFileName();
            XpsDocument exportdoc = null;

            try
            {
                doc.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypeXPS, tempoutxpsname, PbFixedFormatIntent.pbIntentPrinting);
                Thread.Sleep(100);
                File.Copy(tempoutxpsname, tempcopyxpsname, true);
                exportdoc = new XpsDocument(tempcopyxpsname, FileAccess.Read);
            }
            finally
            {
                File.Delete(tempoutxpsname);
            }

            return exportdoc;
        }

        public void SavePUB(string savepath){
            string outbasename = FileNameEscape(curname);
            string outpubname = Path.GetFullPath(savepath) + Path.DirectorySeparatorChar + outbasename + ".pub";

            File.Copy(doc.FullName, outpubname, true);
        }

        public ReportCardData DataSource
        {
            get
            {
                return this.data;
            }
            set
            {
                this.data = value;
            }
        }

        public string DataSourceName
        {
            get
            {
                return datasourcename;
            }
        }

        public string DataFilter
        {
            get
            {
                return datafilter;
            }
        }

        public string[] Names
        {
            get
            {
                if (deleted)
                {
                    throw new ObjectDisposedException("ReportCard");
                }
                return data.Names;
            }
        }

        public PageSetup PageSetup
        {
            get
            {
                return this.doc.PageSetup;
            }
        }

        public static ReportCard OpenTemplate(string pubname, Publisher.Application pubapp, bool usewingdingticks)
        {
            bool pubappstarted = false;
            string temppubname = null;
            Document doc = null;
            ReportCard card = null;
            bool done = false;

            try
            {
                if (pubapp == null)
                {
                    pubapp = new Publisher.Application();
                    pubappstarted = true;
                }
                if (!File.Exists(pubname))
                {
                    throw new FileNotFoundException("File not found", pubname);
                }
                temppubname = Path.GetTempFileName();
                File.Copy(pubname, temppubname, true);
                doc = pubapp.Open(temppubname, false, false, PbSaveOptions.pbDoNotSaveChanges);
                card = new ReportCard((pubappstarted ? pubapp : null), doc, pubname, usewingdingticks);
                done = true;
                return card;
            }
            finally
            {
                if (!done)
                {
                    if (card != null)
                    {
                        card.Dispose();
                        card = null;
                        doc = null;
                        pubapp = null;
                    }

                    if (pubappstarted)
                    {
                        if (doc != null)
                        {
                            doc.Close();
                            doc = null;
                        }

                        ((_Application)pubapp).Quit();
                        pubapp = null;
                    }
                }
            }
        }        
    }

    public enum ReportCardWorkerMessage
    {
        Noop,
        Quit,
        OpenDatasource,
        CloseDatasource,
        OpenTemplate,
        CloseTemplate,
        MergeRecord,
        SavePUB,
        SavePDF,
        ExportXPS,
        GetPageSetup
    }

    internal delegate void ReportCardWorkerJobHandler(object sender, ReportCardWorkerJob e);

    internal class ReportCardWorkerJob : EventArgs
    {
        public ReportCardWorkerMessage Msg;
        public string Name;
        public object Data;
        public Exception Error;
        public ReportCardWorkerJobHandler Completed;
        public bool Cancel;
        public bool Retry;

        public ReportCardWorkerJob()
        {
            Msg = ReportCardWorkerMessage.Noop;
            Name = null;
            Data = null;
            Error = null;
            Completed = null;
            Cancel = false;
        }
    }

    internal class ReportCardWorker : IDisposable
    {
        private Queue jobqueue;
        private AutoResetEvent jobadded;
        private Thread worker;
        private ReportCard card;
        private ReportCardData data;
        private Object cancelled;
        private bool usewingdingticks;
        private Publisher.Application pubapp;
        private Object pubapplock = new object();

        public ReportCardWorker()
        {
            jobqueue = new Queue();
            jobadded = new AutoResetEvent(false);
            cancelled = false;
            worker = new Thread(Run);
            worker.Start();
        }

        ~ReportCardWorker()
        {
            this.Dispose(false);
        }

        public void Dispose()
        {
            this.Dispose(true);
        }

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.Quit();
                this.worker.Join();
            }

            if (jobadded != null)
            {
                this.jobadded.Close();
                this.jobadded = null;
            }

            if (this.card != null)
            {
                this.card.Dispose();
                this.card = null;
            }

            lock (this.pubapplock)
            {
                if (this.pubapp != null)
                {
                    ((_Application)this.pubapp).Quit();
                    this.pubapp = null;
                }
            }

            if (this.data != null)
            {
                this.data.Dispose();
                this.data = null;
            }

            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

        protected void Run()
        {
            ReportCardWorkerJob job;

            pubapp = new Publisher.Application();

            while (jobadded.WaitOne())
            {
                int count;
                lock (cancelled)
                {
                    cancelled = false;
                }
                lock (jobqueue.SyncRoot)
                {
                    count = jobqueue.Count;
                }
                while (count != 0)
                {
                    lock (jobqueue.SyncRoot)
                    {
                        if (jobqueue.Count == 0)
                        {
                            break;
                        }
                        job = (ReportCardWorkerJob)jobqueue.Dequeue();
                        count = jobqueue.Count;
                    }

                    do
                    {
                        job.Retry = false;
                        job.Cancel = false;
                        job.Error = null;
                        try
                        {
                            Console.WriteLine("Processing job {0}\n", job.Msg.ToString());
                            switch (job.Msg)
                            {
                                case ReportCardWorkerMessage.OpenDatasource:
                                    if (card != null)
                                    {
                                        card.DataSource = null;
                                    }
                                    if (data != null)
                                    {
                                        data.Dispose();
                                    }
                                    data = new ReportCardData(job.Name, DataFilter);
                                    job.Data = data.Names;
                                    if (card != null)
                                    {
                                        card.DataSource = data;
                                    }
                                    break;
                                case ReportCardWorkerMessage.CloseDatasource:
                                    if (card != null)
                                    {
                                        card.DataSource = null;
                                    }
                                    if (data != null)
                                    {
                                        data.Dispose();
                                        data = null;
                                    }
                                    break;
                                case ReportCardWorkerMessage.OpenTemplate:
                                    if (card != null)
                                    {
                                        card.Dispose();
                                    }
                                    card = ReportCard.OpenTemplate(job.Name, pubapp, usewingdingticks);
                                    if (data != null)
                                    {
                                        card.DataSource = data;
                                    }
                                    break;
                                case ReportCardWorkerMessage.CloseTemplate:
                                    if (card != null)
                                    {
                                        card.Dispose();
                                        card = null;
                                    }
                                    break;
                                case ReportCardWorkerMessage.MergeRecord:
                                    card.MergeReport(job.Name);
                                    break;
                                case ReportCardWorkerMessage.SavePUB:
                                    card.SavePUB(job.Name);
                                    break;
                                case ReportCardWorkerMessage.SavePDF:
                                    card.SavePDF(job.Name);
                                    break;
                                case ReportCardWorkerMessage.ExportXPS:
                                    job.Data = card.ExportXPS();
                                    break;
                                case ReportCardWorkerMessage.GetPageSetup:
                                    job.Data = card.PageSetup;
                                    break;
                                case ReportCardWorkerMessage.Quit:
                                    lock (pubapplock)
                                    {
                                        if (pubapp != null)
                                        {
                                            ((Publisher._Application)pubapp).Quit();
                                        }
                                        pubapp = null;
                                    }
                                    return;
                            }
                        }
                        catch (Exception e)
                        {
                            job.Error = e;
                            /*
                            Console.Write(
                                "Caught exception\n" + 
                                "Msg = {0}\n" +
                                "Name = {1}\n" +
                                "Data = {2}\n" +
                                "{3}\n\n",
                                job.Msg.ToString(),
                                job.Name,
                                job.Data.ToString(),
                                e.ToString());
                             */
                        }
                        finally
                        {
                            if (job.Completed != null && !(bool)cancelled)
                            {
                                job.Retry = false;
                                job.Cancel = false;

                                job.Completed(this, job);
                                if (job.Cancel)
                                {
                                    job.Retry = false;
                                    jobqueue.Clear();
                                }
                            }
                            cancelled = false;
                        }
                    } while (job.Retry);
                }
            }
        }

        public void Enqueue(ReportCardWorkerJob job)
        {
            lock (jobqueue.SyncRoot)
            {
                /*
                if (job.Msg != ReportCardWorkerMessage.Noop)
                {
                    foreach (ReportCardWorkerJob curjob in jobqueue)
                    {
                        if (curjob.Msg != ReportCardWorkerMessage.Noop &&
                            curjob.Msg == job.Msg &&
                            curjob.Name == job.Name &&
                            curjob.Data == job.Data)
                        {
                            if (curjob.Completed != job.Completed)
                            {
                                curjob.Completed += job.Completed;
                            }
                            return;
                        }
                    }
                }
                 */

                jobqueue.Enqueue(job);
            }
            jobadded.Set();
        }

        public void Enqueue(ReportCardWorkerMessage msg, string name, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(new ReportCardWorkerJob { Msg = msg, Name = name, Completed = callback, Data = data });
        }

        public void BeginNoop(ReportCardWorkerJobHandler callback, string name, object data)
        {
            Enqueue(ReportCardWorkerMessage.Noop, name, callback, data);
        }

        public void BeginOpenTemplate(string name, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.OpenTemplate, name, callback, data);
        }

        public void BeginCloseTemplate(ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.CloseTemplate, null, callback, data);
        }

        public void BeginOpenDatasource(string name, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.OpenDatasource, name, callback, data);
        }

        public void BeginCloseDatasource(ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.CloseDatasource, null, callback, data);
        }

        public void BeginMergeRecord(string name, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.MergeRecord, name, callback, data);
        }

        public void BeginSavePUB(string path, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.SavePUB, path, callback, data);
        }

        public void BeginSavePDF(string path, ReportCardWorkerJobHandler callback, object data)
        {
            Enqueue(ReportCardWorkerMessage.SavePDF, path, callback, data);
        }

        public void BeginExportXPS(ReportCardWorkerJobHandler callback)
        {
            Enqueue(ReportCardWorkerMessage.ExportXPS, null, callback, null);
        }

        public void BeginGetPageSetup(ReportCardWorkerJobHandler callback)
        {
            Enqueue(ReportCardWorkerMessage.GetPageSetup, null, callback, null);
        }

        public void Quit()
        {
            CancelAll();

            Enqueue(ReportCardWorkerMessage.CloseTemplate, null, null, null);
            Enqueue(ReportCardWorkerMessage.CloseDatasource, null, null, null);
            Enqueue(ReportCardWorkerMessage.Quit, null, null, null);
        }

        public void CancelAll()
        {
            lock (jobqueue.SyncRoot)
            {
                jobqueue.Clear();
            }
            lock (cancelled)
            {
                cancelled = true;
            }
        }

        public string DataSourcePath
        {
            get
            {
                if (card != null)
                {
                    return card.DataSourceName;
                }
                else
                {
                    return null;
                }
            }
        }

        public string DataFilter
        {
            get
            {
                if (card != null)
                {
                    return card.DataFilter;
                }
                else
                {
                    return null;
                }
            }
        }

        public bool UseWingdingTicks
        {
            get
            {
                return usewingdingticks;
            }
            set
            {
                usewingdingticks = value;
            }
        }

        public IEnumerable<DataRow> Data
        {
            get
            {
                return data.AsEnumerable();
            }
        }
    }
}

