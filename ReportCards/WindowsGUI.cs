using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using SouthernCluster.ReportCards;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Printing;
using Publisher = Microsoft.Office.Interop.Publisher;

namespace SouthernCluster.ReportCards
{
    public partial class ReportCardWindowsGUIMerger : Form, IDisposable
    {
        private string[] names;
        private IList<string> initialnames;
        private string templatepath;
        private string datasourcepath;
        private string savetopath;
        private bool exportpdf;
        private bool exportpub;
        private bool usewingdingticks;
        private ReportCardWorker worker;
        private bool templategood;
        private bool datasourcegood;
        private bool savetogood;
        private System.Windows.Controls.PrintDialog pdPrint;

        private static void Usage()
        {
            Console.Out.WriteLine(
                "Usage: {0} [/template:<Template> [/datasource:<Datasource>]\n" +
                "             [/saveto:<SaveToDir>] [/[no-]pdf] [/[no-]pub]\n" +
                "             [/record:<Record>] ]\n" +
                "\n" +
                "  /template:<Template>\n" +
                "        Specify the template to use.\n" +
                "  /datasource:<DataSource>\n" +
                "        Specify the excel workbook to use.\n" +
                "        If this is not specified, it will be based on the template\n" +
                "          name\n" +
                "  /saveto:<SaveToDir>\n" +
                "        Specify the directory to save the exported documents to\n" +
                "        If this is not specified, exported documents will be saved\n" +
                "          in the directory the template is in\n" +
                "  /pdf\n" +
                "  /no-pdf\n" +
                "        Export or don't export PDF documents for each record\n" +
                "  /pub\n" +
                "  /no-pub\n" +
                "        Export or don't export Publisher documents for each record\n" +
                "  /wingdings\n" +
                "  /no-wingdings\n" +
                "        Use or don't use Wingding Check marks for ticks\n" +
                "  /record:<Record>\n" +
                "        Specify one or more records to export\n" +
                "        If this is not specified, all records will be exported\n" +
                "\n",
                Environment.GetCommandLineArgs()[0]);
        }

        public class Options
        {
            public bool ExportPdf = true;
            public bool ExportPub = false;
            public bool UseWingdingTicks = true;
            public string TemplateName = null;
            public string DataSourceName = null;
            public string SaveDir = null;
            public List<string> Names = new List<string>();
        }

        private static bool ParseOptions(string progdir, string[] args, out Options options)
        {
            options = new Options();
            bool endnamed = false;
            string optname = null;
            bool help = false;
            Queue<string> unnamed = new Queue<string>();

            for (int i = 0; i < args.Length; i++)
            {
                string optarg = null;

                if (args[i].StartsWith("--"))
                {
                    optname = args[i].Substring(2);
                }
                else if (args[i].StartsWith("-") || args[i].StartsWith("/"))
                {
                    optname = args[i].Substring(1);
                }
                else
                {
                    optarg = args[i];
                }

                if (optname == "")
                {
                    endnamed = true;
                }

                if (!endnamed && optname != null)
                {
                    if (optname.Contains(":"))
                    {
                        optarg = optname.Substring(optname.IndexOf(":") + 1);
                        optname = optname.Substring(0, optname.IndexOf(":")).ToLower();
                    }
                    if (optname != null &&
                        optname.Length >= 1 &&
                        ("template".StartsWith(optname) ||
                         "datasource".StartsWith(optname) ||
                         "savedir".StartsWith(optname) ||
                         "record".StartsWith(optname)))
                    {

                        if (optarg == null && i + 1 < args.Length)
                        {
                            i++;
                            optarg = args[i];
                        }

                        if ("template".StartsWith(optname))
                        {
                            optname = "template";
                        }
                        else if ("datasource".StartsWith(optname))
                        {
                            optname = "datasource";
                        }
                        else if ("savedir".StartsWith(optname))
                        {
                            optname = "savedir";
                        }
                        else if ("record".StartsWith(optname))
                        {
                            optname = "record";
                        }
                    }

                    switch (optname)
                    {
                        case "pdf":
                            options.ExportPdf = true;
                            optname = null;
                            break;
                        case "no-pdf":
                            options.ExportPdf = false;
                            optname = null;
                            break;
                        case "pub":
                            options.ExportPub = true;
                            optname = null;
                            break;
                        case "no-pub":
                            options.ExportPub = false;
                            optname = null;
                            break;
                        case "wingdings":
                            options.UseWingdingTicks = true;
                            optname = null;
                            break;
                        case "no-wingdings":
                            options.UseWingdingTicks = false;
                            optname = null;
                            break;
                        case "template":
                            try
                            {
                                options.TemplateName = Path.GetFullPath(optarg);
                            }
                            catch
                            {
                                options.TemplateName = optarg;
                            }

                            optname = null;
                            break;
                        case "datasource":
                            try
                            {
                                options.DataSourceName = Path.GetFullPath(optarg);
                            }
                            catch
                            {
                                options.DataSourceName = optarg;
                            }

                            optname = null;
                            break;
                        case "savedir":
                            options.SaveDir = Path.GetFullPath(optarg);
                            optname = null;
                            break;
                        case "record":
                            options.Names.Add(optarg);
                            break;
                        case "help":
                            help = true;
                            break;
                        case "?":
                            goto case "help";
                        default:
                            goto case "help";
                    }
                }
                else
                {
                    unnamed.Enqueue(optarg);
                }
            }

            if (options.TemplateName == null && unnamed.Count != 0)
            {
                string name = unnamed.Peek();
                if (name.ToLower().EndsWith(".pub"))
                {
                    name = unnamed.Dequeue();
                }
                else if (name.ToLower().EndsWith(".xls") ||
                         name.ToLower().EndsWith(".xlsx") ||
                         name.ToLower().EndsWith(".xlsb") ||
                         name.ToLower().EndsWith(".xlsm"))
                {
                    options.DataSourceName = unnamed.Dequeue();
                    name = options.DataSourceName.Substring(0, options.DataSourceName.LastIndexOf('.')) + ".pub";
                }
                else
                {
                    name = null;
                    help = true;
                }

                try
                {
                    if (name != null)
                    {
                        options.TemplateName = Path.GetFullPath(name);
                    }
                }
                catch
                {
                    options.TemplateName = name;
                }
            }

            if (unnamed.Count != 0)
            {
                options.Names.AddRange(unnamed);
                unnamed.Clear();
            }

            if (String.IsNullOrEmpty(options.SaveDir))
            {
                if (options.TemplateName != null)
                {
                    options.SaveDir = Path.GetDirectoryName(Path.GetFullPath(options.TemplateName));
                }
                else
                {
                    options.SaveDir = progdir;
                }
            }

            return !help;
        }

        public static int Main(string[] args)
        {
            bool help;
            string progname;
            Options options;

            try
            {
                progname = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
                //progname = Environment.GetCommandLineArgs()[0];
            }
            catch
            {
                MessageBox.Show("How can I run if I don't even know where I am?\n" +
                                "Please copy this program, the report to be merged and the excel datasource to be used into a folder on your C: drive.\n" +
                                "If this problem persists, please contact your system administrator.\n" +
                                "If your system administrator cannot resolve this, please contact Ben Peddell on 0431 800 195.");
                return 1;
            }

            string progdir = Path.GetDirectoryName(progname);
            help = !ParseOptions(progdir, args, out options);

            if (args != null && args.Length != 0 && NativeMethods.TryAttachConsole(-1))
            {
                ReportCard card = null;
                ReportCardData data = null;

                if (help)
                {
                    Usage();
                    return 1;
                }

                try
                {
                    Console.Out.Write("Initializing report template and datasource ... ");
                    card = ReportCard.OpenTemplate(options.TemplateName, null, options.UseWingdingTicks);
                    if (options.DataSourceName == null)
                    {
                        options.DataSourceName = card.DataSourceName;
                        if (options.DataSourceName == null)
                        {
                            Console.Error.WriteLine("Unable to find datasource for template");
                            return 1;
                        }
                    }
                    using (data = new ReportCardData(options.DataSourceName))
                    {
                        Console.Out.Write("Done\n");
                        if (options.Names.Count == 0)
                        {
                            options.Names.AddRange(card.Names);
                        }
                        foreach (string name in options.Names)
                        {
                            Console.Out.Write("Merging record for {0} ... ", name);
                            card.MergeReport(name);
                            if (options.ExportPub)
                            {
                                Console.Out.Write("PUB ... ");
                                card.SavePUB(options.SaveDir);
                            }
                            if (options.ExportPdf)
                            {
                                Console.Out.Write("PDF ... ");
                                card.SavePDF(options.SaveDir);
                            }
                            Console.Out.Write("Done\n");
                        }
                        Console.Out.Write("\n");
                    }
                    return 0;
                }
                catch (Exception e)
                {
                    Console.Out.Write("Failed\n\n");
                    Usage();
                    Console.Out.WriteLine("Caught exception {0}\n{1}\n", e.GetType().Name, e.ToString());
                    return 1;
                }
                finally
                {
                    if (card != null)
                    {
                        card.Dispose();
                    }
                }
            }
            else
            {
                ReportCardWindowsGUIMerger.StartMerge(options);
                return 0;
            }
        }

        public static void StartMerge(Options options)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ReportCardWindowsGUIMerger merger = new ReportCardWindowsGUIMerger(options);
            Application.Run(merger);
        }

        public ReportCardWindowsGUIMerger(Options options)
        {
            InitializeComponent();
            this.templategood = false;
            this.datasourcegood = false;
            this.templatepath = options.TemplateName;
            this.datasourcepath = options.DataSourceName;
            this.savetopath = options.SaveDir;
            this.savetogood = Directory.Exists(options.SaveDir);
            this.exportpdf = options.ExportPdf;
            this.exportpub = options.ExportPub;
            this.usewingdingticks = options.UseWingdingTicks;
            this.initialnames = options.Names;
            this.pdPrint = new System.Windows.Controls.PrintDialog();
            this.worker = new ReportCardWorker();
            this.worker.UseWingdingTicks = options.UseWingdingTicks;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }

            if (this.worker != null)
            {
                this.worker.Dispose();
                this.worker = null;
            }

            base.Dispose(disposing);
        }

        private void BrowseForTemplate()
        {
            templategood = false;
            datasourcegood = false;

            if (tbTemplate.Text != "")
            {
                string path = tbTemplate.Text;
                while (path != null && !File.Exists(path) && !Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                }

                if (path != null)
                {
                    if (Directory.Exists(path))
                    {
                        fdTemplateOpen.InitialDirectory = path;
                    }
                    else if (File.Exists(path))
                    {
                        fdTemplateOpen.InitialDirectory = Path.GetDirectoryName(path);
                        fdTemplateOpen.FileName = path;
                    }
                }
                else
                {
                    fdTemplateOpen.InitialDirectory = Application.StartupPath;
                }
            }
            else
            {
                fdTemplateOpen.InitialDirectory = Application.StartupPath;
            }

            if (fdTemplateOpen.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                templatepath = fdTemplateOpen.FileName;
                tbTemplate.Text = templatepath;
                AutoFillSaveto();
                OpenTemplate();
            }
        }

        private void BrowseForDatasource()
        {
            datasourcegood = false;

            if (!String.IsNullOrEmpty(tbDatasource.Text))
            {
                string path = tbDatasource.Text;
                while (!String.IsNullOrEmpty(path) && !File.Exists(path) && !Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                }
                if (String.IsNullOrEmpty(path))
                {
                    path = Path.GetDirectoryName(tbTemplate.Text);
                    while (!String.IsNullOrEmpty(path) && !File.Exists(path) && !Directory.Exists(path))
                    {
                        path = Path.GetDirectoryName(path);
                    }
                }

                if (!String.IsNullOrEmpty(path))
                {
                    if (Directory.Exists(path))
                    {
                        fdDatasourceOpen.InitialDirectory = path;
                    }
                    else if (File.Exists(path))
                    {
                        fdDatasourceOpen.InitialDirectory = Path.GetDirectoryName(path);
                        fdDatasourceOpen.FileName = path;
                    }
                }
                else
                {
                    fdDatasourceOpen.InitialDirectory = Application.StartupPath;
                }
            }
            else
            {
                fdDatasourceOpen.InitialDirectory = Application.StartupPath;
            }

            if (fdDatasourceOpen.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                datasourcepath = fdDatasourceOpen.FileName;
                tbDatasource.Text = datasourcepath;
                OpenDatasource();
            }
        }

        private void BrowseForSaveto()
        {
            if (!String.IsNullOrEmpty(tbSaveTo.Text))
            {
                string path = tbSaveTo.Text;
                while (!String.IsNullOrEmpty(path) && !Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                }

                if (!String.IsNullOrEmpty(path))
                {
                    if (Directory.Exists(path))
                    {
                        fdSaveTo.SelectedPath = path;
                    }
                }
                else
                {
                    fdSaveTo.SelectedPath = Application.StartupPath;
                }
            }
            else
            {
                fdSaveTo.SelectedPath = Application.StartupPath;
            }

            if (fdSaveTo.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                savetopath = fdSaveTo.SelectedPath;
                tbSaveTo.Text = savetopath;
                tbSaveTo.ScrollToCaret();
                savetogood = true;
                if (datasourcegood && templategood)
                {
                    btnMerge.Enabled = true;
                }
            }
        }

        private void DisableControls()
        {
            grpMergeFrom.Enabled = false;
            grpMergeType.Enabled = false;
            btnMerge.Enabled = false;
            btnPrint.Enabled = false;
            btnCancel.Enabled = false;
            grpNames.Enabled = false;
        }

        private void DisableControlsCancellable()
        {
            grpMergeFrom.Enabled = false;
            grpMergeType.Enabled = false;
            btnMerge.Enabled = false;
            btnPrint.Enabled = false;
            btnCancel.Enabled = true;
            grpNames.Enabled = false;
        }

        private void EnableControls()
        {
            grpMergeFrom.Enabled = true;
            grpMergeType.Enabled = true;
            btnMerge.Enabled = true;
            btnPrint.Enabled = true;
            btnCancel.Enabled = false;
            grpNames.Enabled = true;
        }

        private void AutoFillDatasource() /* DELETE */
        {
            string path = templatepath;
            string basename;

            if (path.EndsWith(".pub"))
            {
                basename = path.Substring(0, path.Length - 4);
            }
            else
            {
                basename = path;
            }

            path = null;

            if (File.Exists(basename + ".xlsx"))
            {
                path = basename + ".xlsx";
            }
            else if (File.Exists(basename + ".xls"))
            {
                path = basename + ".xls";
            }
            else if (File.Exists(basename + ".xlsb"))
            {
                path = basename + ".xlsb";
            }

            if (path != null)
            {
                datasourcepath = path;
                tbDatasource.Text = path;
            }
        }

        private void AutoFillSaveto()
        {
            string path = Path.GetDirectoryName(templatepath);
            if (Directory.Exists(path))
            {
                savetopath = path;
                tbSaveTo.Text = savetopath;
                savetogood = true;
            }
        }

        private void OpenDatasource()
        {
            if (datasourcepath != "")
            {
                DisableControls();
                worker.BeginCloseTemplate(null, null);
                worker.BeginCloseDatasource(null, null);
                templategood = false;
                datasourcegood = false;
                ssStatusText.Text = "Opening datasource";
                worker.BeginOpenDatasource(datasourcepath, SetNames, null);
            }
        }

        private void OpenTemplate()
        {
            if (templatepath != "" && File.Exists(templatepath))
            {
                DisableControls();
                worker.BeginCloseTemplate(null, null);
                worker.BeginCloseDatasource(null, null);
                templategood = false;
                datasourcegood = false;
                btnMerge.Enabled = false;
                ssStatusText.Text = "Opening template";
                worker.BeginOpenTemplate(templatepath, TemplateOpened, null);
            }
        }

        private void MergeReports(bool print)
        {
            if (datasourcegood && templategood && savetogood)
            {
                int numselected = clbNames.CheckedItems.Count;
                if (cbMergeToPUB.CheckState == CheckState.Unchecked &&
                    cbMergeToPDF.CheckState == CheckState.Unchecked &&
                    print == false)
                {
                    MessageBox.Show("Please select PDF and/or PUB export\n", "No export formats selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (numselected == 0)
                {
                    MessageBox.Show("Please select at least one record\n", "No records selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DisableControlsCancellable();

                    string[] selnames = new string[numselected];
                    int index = 0;
                    foreach (object o in clbNames.CheckedItems)
                    {
                        selnames[index++] = (string)o;
                    }

                    for (int i = 0; i < numselected; i++)
                    {
                        string name = selnames[i];
                        worker.BeginNoop(UpdateStatusBar, "Merging document for " + name, (i * 100 + 33) / numselected);
                        worker.BeginMergeRecord(name, OperationStatus, name);
                        if (print)
                        {
                            worker.BeginNoop(UpdateStatusBar, "Printing report for " + name, (i * 100 + 100) / numselected);
                            worker.BeginExportXPS(PrintXPSDocument);
                        }
                        else
                        {
                            if (cbMergeToPUB.CheckState != CheckState.Unchecked)
                            {
                                worker.BeginNoop(UpdateStatusBar, "Exporting PUB for " + name, (i * 100 + 66) / numselected);
                                worker.BeginSavePUB(savetopath, OperationStatus, name);
                            }
                            if (cbMergeToPDF.CheckState != CheckState.Unchecked)
                            {
                                worker.BeginNoop(UpdateStatusBar, "Exporting PDF for " + name, (i * 100 + 100) / numselected);
                                worker.BeginSavePDF(savetopath, OperationStatus, name);
                            }
                        }
                    }
                    worker.BeginNoop(MergeComplete, null, null);
                }
            }
        }

        private void CancelMerge()
        {
            worker.CancelAll();
            ssProgress.Visible = false;
            ssStatusText.Text = "Merge cancelled";
            EnableControls();
        }

        private void MergeComplete(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(MergeComplete), new object[] { sender, e });
            }
            else
            {
                ssProgress.Visible = false;
                ssStatusText.Text = "Merge complete";
                EnableControls();
            }
        }

        private void OperationStatus(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(OperationStatus), new object[] { sender, e });
            }
            else
            {
                if (e.Error != null)
                {
                    string msg;
                    if (e.Error is System.IO.IOException && e.Msg == ReportCardWorkerMessage.SavePDF)
                    {
                        msg = String.Format(
                            "Could not write to the PDF file when merging record {0}\n" +
                            "Please check that the PDF file is not open\n" +
                            "Do you wish to retry the operation, skip the record, or cancel?",
                            e.Data.ToString()
                        );
                    }
                    else
                    {
                        msg = String.Format(
                            "Caught error when processing record {0}\n" +
                            "Do you wish to retry the operation, skip the record, or cancel?\n\n" +
                            "  Operation = {1}\n" +
                            "  Name = {2}\n" +
                            "  Error:\n{3}\n",
                            e.Data.ToString(),
                            e.Msg.ToString(),
                            e.Name,
                            e.Error.ToString()
                        );
                    }

                    int res = NativeMethods.MsgBox(
                        IntPtr.Zero,
                        msg,
                        "Error processing record",
                        NativeMethods.MB_CANCELTRYCONTINUE |
                        NativeMethods.MB_ICONERROR |
                        NativeMethods.MB_TASKMODAL |
                        NativeMethods.MB_DEFBUTTON2
                    );
                    if (res == NativeMethods.IDCANCEL)
                    {
                        e.Cancel = true;
                        ssStatusText.Text = "Error processing record";
                        CancelMerge();
                    }
                    else if (res == NativeMethods.IDTRYAGAIN)
                    {
                        e.Retry = true;
                    }
                }
            }
        }

        private void PrintXPSDocument(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(PrintXPSDocument), new object[] { sender, e });
            }
            else
            {
                XpsDocument exps = e.Data as XpsDocument;
                if (exps != null)
                {
                    try
                    {
                        pdPrint.PrintDocument(exps.GetFixedDocumentSequence().DocumentPaginator, Path.GetFileNameWithoutExtension(this.templatepath));
                    }
                    catch
                    {
                        CancelMerge();
                    }
                }
            }
        }

        private void UpdateStatusBar(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(UpdateStatusBar), new object[] { sender, e });
            }
            else
            {
                ssStatusText.Text = e.Name;
                if (e.Data == null)
                {
                    ssProgress.Visible = false;
                }
                else if (e.Data is int)
                {
                    ssProgress.Value = (int)e.Data;
                    ssProgress.Visible = true;
                }
                else if (e.Data is float)
                {
                    ssProgress.Value = (int)((float)e.Data * 100);
                    ssProgress.Visible = true;
                }
                else
                {
                    ssProgress.Visible = false;
                }
            }
        }

        private void SetNames(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(SetNames), new object[] { sender, e });
            }
            else
            {
                if (e.Error != null)
                {
                    ssStatusText.Text = "Could not open datasource";
                    MessageBox.Show("Could not open the specified datasource.\n" +
                                    "Please specify the Microsoft Excel workbook to use as a datasource.\n" +
                                    e.Error.ToString(), "Could not open datasource", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnDatasourceBrowse_Click(this, EventArgs.Empty);
                }
                else
                {
                    this.names = (string[])e.Data;
                    clbNames.Items.Clear();
                    foreach (string name in names)
                    {
                        clbNames.Items.Add(name);
                        if (this.initialnames == null || this.initialnames.Count == 0 || this.initialnames.Contains(name))
                        {
                            clbNames.SetItemCheckState(clbNames.Items.IndexOf(name), CheckState.Checked);
                        }
                    }
                    this.initialnames = null;

                    datasourcegood = true;
                    ssStatusText.Text = "Ready";
                    EnableControls();
                }
            }
        }

        private void TemplateOpened(object sender, ReportCardWorkerJob e)
        {
            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(TemplateOpened), new object[] { sender, e });
            }
            else
            {
                if (e.Error != null)
                {
                    ssStatusText.Text = "Could not open template";
                    MessageBox.Show("Could not open the specified template.\n" +
                                    "Please specify the Microsoft Publisher template to use.\n" +
                                    e.Error.ToString(), "Could not open template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnTemplateBrowse_Click(this, EventArgs.Empty);
                }
                else
                {
                    ssStatusText.Text = "Fetching Page Setup";
                    worker.BeginGetPageSetup(PageSetup);
                }
            }
        }

        private void PageSetup(object sender, ReportCardWorkerJob e)
        {
            Publisher.PageSetup setup = (Publisher.PageSetup)e.Data;
            double width = (float)setup.PageWidth * 96.0 / 72.0;
            double height = (float)setup.PageHeight * 96.0 / 72.0;

            if (InvokeRequired)
            {
                Invoke(new ReportCardWorkerJobHandler(PageSetup), new object[] { sender, e });
            }
            else
            {
                pdPrint.UserPageRangeEnabled = false;

                pdPrint.PrintTicket.PageMediaSize = new PageMediaSize(width, height);
                pdPrint.PrintTicket.CopyCount = 1;
                pdPrint.PrintTicket.Collation = Collation.Collated;
                pdPrint.PrintTicket.PageMediaType = PageMediaType.Plain;
                if (width > height)
                {
                    pdPrint.PrintTicket.PageOrientation = PageOrientation.Landscape;
                }
                else
                {
                    pdPrint.PrintTicket.PageOrientation = PageOrientation.Portrait;
                }

                templategood = true;

                if (worker.DataSourcePath != null)
                {
                    datasourcepath = worker.DataSourcePath;
                    tbDatasource.Text = datasourcepath;
                }

                if (datasourcepath != null)
                {
                    OpenDatasource();
                }
                else
                {
                    btnDatasourceBrowse_Click(this, EventArgs.Empty);
                }
            }
        }

        private void ReportCardWindowsGUIMerger_Load(object sender, EventArgs e)
        {
            btnMerge.Enabled = false;
            btnPrint.Enabled = false;
            btnCancel.Enabled = false;
            ssProgress.Visible = false;
            ssStatusText.Text = "Ready";
            cbMergeToPDF.Checked = exportpdf;
            cbMergeToPUB.Checked = exportpub;
            if (savetopath != null)
            {
                tbSaveTo.Text = savetopath;
            }
            if (templatepath != null)
            {
                tbTemplate.Text = templatepath;
                if (savetopath == null)
                {
                    AutoFillSaveto();
                }
                OpenTemplate();
            }
        }

        private void btnTemplateBrowse_Click(object sender, EventArgs e)
        {
            BrowseForTemplate();
        }

        private void btnDatasourceBrowse_Click(object sender, EventArgs e)
        {
            BrowseForDatasource();
        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
            MergeReports(false);
        }

        private void btnSavetoBrowse_Click(object sender, EventArgs e)
        {
            BrowseForSaveto();
        }

        private void tbTemplate_TextChanged(object sender, EventArgs e)
        {
            if (tbTemplate.Text != templatepath)
            {
                worker.CancelAll();
                templatepath = tbTemplate.Text;
                templategood = false;
                datasourcegood = false;
                btnMerge.Enabled = false;
                btnPrint.Enabled = false;

                if (templatepath != "" && File.Exists(templatepath))
                {
                    AutoFillSaveto();
                    OpenDatasource();
                }
            }
        }

        private void tbDatasource_TextChanged(object sender, EventArgs e)
        {
            if (tbDatasource.Text != datasourcepath)
            {
                worker.CancelAll();
                datasourcepath = tbDatasource.Text;
                datasourcegood = false;
                btnMerge.Enabled = false;
                btnPrint.Enabled = false;

                if (datasourcepath != "" && File.Exists(datasourcepath))
                {
                    OpenDatasource();
                }
            }
        }

        private void tbSaveTo_TextChanged(object sender, EventArgs e)
        {
            if (tbSaveTo.Text != savetopath)
            {
                savetogood = false;
                savetopath = tbSaveTo.Text;
                btnMerge.Enabled = false;
                btnPrint.Enabled = false;

                if (savetopath != "" && Directory.Exists(savetopath))
                {
                    savetogood = true;
                    if (datasourcegood && templategood)
                    {
                        btnMerge.Enabled = true;
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            CancelMerge();
        }

        private void cbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (cbSelectAll.CheckState != CheckState.Indeterminate)
            {
                CheckState state = cbSelectAll.CheckState;
                for (int i = 0; i < clbNames.Items.Count; i++)
                {
                    if (clbNames.GetItemCheckState(i) != state)
                    {
                        clbNames.SetItemCheckState(i, state);
                    }
                }
            }
        }

        private void clbNames_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue != cbSelectAll.CheckState)
            {
                cbSelectAll.CheckState = CheckState.Indeterminate;
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (pdPrint.ShowDialog() == true)
            {
                MergeReports(true);
            }
        }
    }
}
