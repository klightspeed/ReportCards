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

public partial class ReportCardWindowsGUIMerger : Form
{
    private string[] names;
    private List<string> initialnames;
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

    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(int pid);

    private static bool TryAttachConsole(int pid)
    {
        try
        {
            return AttachConsole(pid);
        }
        catch (System.Security.SecurityException)
        {
            return false;
        }
    }
    
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

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            int ret = RealMain(args);
            Console.Out.Flush();
            Console.Error.Flush();
            return ret;
        }
        catch (Exception e)
        {
            MessageBox.Show("Caught exception:\n" + e.ToString());
            return 1;
        }
    }

    private static int RealMain(string[] args)
    {
        string templatename = null;
        string datasourcename = null;
        string savedir = null;
        List<string> names = null;
        bool endnamed = false;
        bool exportpdf = true;
        bool exportpub = false;
        bool usewingdingticks = true;
        string optname = null;
        Queue<string> unnamed = new Queue<string>();
        string progname;
        try
        {
            progname = Environment.GetCommandLineArgs()[0];
        }
        catch
        {
            MessageBox.Show("How can I run if I don't even know where I am?\n" +
                            "Please copy this program and the report to be merged into a folder on your C: drive\n" +
                            "If this problem persists, please contact your system administrator");
            return 1;
        }
        bool help = false;

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
                        exportpdf = true;
                        optname = null;
                        break;
                    case "no-pdf":
                        exportpdf = false;
                        optname = null;
                        break;
                    case "pub":
                        exportpub = true;
                        optname = null;
                        break;
                    case "no-pub":
                        exportpub = false;
                        optname = null;
                        break;
                    case "wingdings":
                        usewingdingticks = true;
                        optname = null;
                        break;
                    case "no-wingdings":
                        usewingdingticks = false;
                        optname = null;
                        break;
                    case "template":
                        templatename = Path.GetFullPath(optarg);
                        optname = null;
                        break;
                    case "datasource":
                        if (!optarg.StartsWith("grubrics:"))
                        {
                            datasourcename = Path.GetFullPath(optarg);
                        }
                        else
                        {
                            datasourcename = optarg;
                        }
                        optname = null;
                        break;
                    case "savedir":
                        savedir = Path.GetFullPath(optarg);
                        optname = null;
                        break;
                    case "record":
                        if (names == null)
                        {
                            names = new List<string>();
                        }
                        names.Add(optarg);
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

        if (templatename == null && unnamed.Count != 0)
        {
            templatename = Path.GetFullPath(unnamed.Dequeue());
        }
        if (names == null && unnamed.Count != 0)
        {
            names = new List<string>(unnamed);
            unnamed.Clear();
        }

        if (args.Length != 0 && TryAttachConsole(-1))
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
                if (savedir == null)
                {
                    savedir = Path.GetDirectoryName(Path.GetFullPath(templatename));
                }

                Console.Out.Write("Initializing report template and datasource ... ");
                card = ReportCard.OpenTemplate(templatename, null, usewingdingticks);
                if (datasourcename == null){
                    datasourcename = card.DataSourceName;
                    if (datasourcename == null){
                        Console.Error.WriteLine("Unable to find datasource for template");
                        return 1;
                    }
                }
                data = new ReportCardData(datasourcename);
                Console.Out.Write("Done\n");
                if (names == null)
                {
                    names = new List<string>(card.Names);
                }
                foreach (string name in names)
                {
                    Console.Out.Write("Merging record for {0} ... ", name);
                    card.MergeReport(name);
                    if (exportpub)
                    {
                        Console.Out.Write("PUB ... ");
                        card.SavePUB(savedir);
                    }
                    if (exportpdf)
                    {
                        Console.Out.Write("PDF ... ");
                        card.SavePDF(savedir);
                    }
                    Console.Out.Write("Done\n");
                }
                Console.Out.Write("\n");
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
            ReportCardWindowsGUIMerger.StartMerge(templatename, datasourcename, savedir, exportpdf, exportpub, names, usewingdingticks);
            return 0;
        }
    }
    
    public static void StartMerge(string templatepath, string datasourcepath, string savetopath, bool exportpdf, bool exportpub, List<string> names, bool usewingdingticks){
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        ReportCardWindowsGUIMerger merger = new ReportCardWindowsGUIMerger(templatepath, datasourcepath, savetopath, exportpdf, exportpub, names, usewingdingticks);
        Application.Run(merger);
    }

    public ReportCardWindowsGUIMerger(string templatepath, string datasourcepath, string savetopath, bool exportpdf, bool exportpub, List<string> names, bool usewingdingticks)
    {
        InitializeComponent();
        this.templategood = false;
        this.datasourcegood = false;
        this.savetogood = false;
        this.templatepath = templatepath;
        this.datasourcepath = datasourcepath;
        this.savetopath = savetopath;
        this.exportpdf = exportpdf;
        this.exportpub = exportpub;
        this.usewingdingticks = usewingdingticks;
        this.initialnames = names;
        this.pdPrint = new System.Windows.Controls.PrintDialog();
        this.worker = new ReportCardWorker();
        this.worker.UseWingdingTicks = usewingdingticks;
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

        if (tbDatasource.Text != "")
        {
            string path = tbDatasource.Text;
            while (path != null && !File.Exists(path) && !Directory.Exists(path))
            {
                path = Path.GetDirectoryName(path);
            }
            if (path == null)
            {
                path = Path.GetDirectoryName(tbTemplate.Text);
                while (path != null && !File.Exists(path) && !Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                }
            }

            if (path != null)
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
        if (tbSaveTo.Text != "")
        {
            string path = tbSaveTo.Text;
            while (path != null && !Directory.Exists(path))
            {
                path = Path.GetDirectoryName(path);
            }

            if (path != null)
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
            ssStatusText.Text = "Opening datasource";
            worker.BeginOpenDatasource(datasourcepath, SetNames, null);
        }
    }

    private void OpenTemplate()
    {
        worker.BeginCloseTemplate(null, null);
        worker.BeginCloseDatasource(null, null);
        templategood = false;
        datasourcegood = false;
        btnMerge.Enabled = false;

        if (templatepath != "" && File.Exists(templatepath))
        {
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
                tbTemplate.Enabled = false;
                btnTemplateBrowse.Enabled = false;
                tbDatasource.Enabled = false;
                btnDatasourceBrowse.Enabled = false;
                tbSaveTo.Enabled = false;
                btnSavetoBrowse.Enabled = false;
                clbNames.Enabled = false;
                btnMerge.Enabled = false;
                btnPrint.Enabled = false;
                cbSelectAll.Enabled = false;
                cbMergeToPDF.Enabled = false;
                cbMergeToPUB.Enabled = false;
                btnCancel.Enabled = true;

                string[] selnames = new string[numselected];
                int index = 0;
                foreach (object o in clbNames.CheckedItems)
                {
                    selnames[index++] = (string)o;
                }

                for (int i = 0; i < numselected; i++){
                    string name = selnames[i];
                    worker.BeginNoop(UpdateStatusBar, "Merging document for " + name, i * 100 / numselected);
                    worker.BeginMergeRecord(name, OperationStatus, name);
                    if (print)
                    {
                        worker.BeginNoop(UpdateStatusBar, "Printing report for " + name, (i * 100 + 50) / numselected);
                        worker.BeginExportXPS(PrintXPSDocument);
                    }
                    else
                    {
                        if (cbMergeToPUB.CheckState != CheckState.Unchecked)
                        {
                            worker.BeginNoop(UpdateStatusBar, "Exporting PUB for " + name, (i * 100 + 25) / numselected);
                            worker.BeginSavePUB(savetopath, OperationStatus, name);
                        }
                        if (cbMergeToPDF.CheckState != CheckState.Unchecked)
                        {
                            worker.BeginNoop(UpdateStatusBar, "Exporting PDF for " + name, (i * 100 + 50) / numselected);
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
        tbTemplate.Enabled = true;
        btnTemplateBrowse.Enabled = true;
        tbDatasource.Enabled = true;
        btnDatasourceBrowse.Enabled = true;
        tbSaveTo.Enabled = true;
        btnSavetoBrowse.Enabled = true;
        clbNames.Enabled = true;
        btnMerge.Enabled = true;
        btnPrint.Enabled = true;
        cbSelectAll.Enabled = true;
        cbMergeToPDF.Enabled = true;
        cbMergeToPUB.Enabled = true;
        btnCancel.Enabled = false;
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
            tbTemplate.Enabled = true;
            btnTemplateBrowse.Enabled = true;
            tbDatasource.Enabled = true;
            btnDatasourceBrowse.Enabled = true;
            tbSaveTo.Enabled = true;
            btnSavetoBrowse.Enabled = true;
            clbNames.Enabled = true;
            btnMerge.Enabled = true;
            btnPrint.Enabled = true;
            cbSelectAll.Enabled = true;
            cbMergeToPDF.Enabled = true;
            cbMergeToPUB.Enabled = true;
            btnCancel.Enabled = false;
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
                DialogResult res = MessageBox.Show(
                    "Caught error when processing record " + (string)e.Data + "\n" +
                    "Do you wish to continue?\n\n" +
                    "  Operation = " + e.Msg.ToString() + "\n" +
                    "  Name = " + e.Name + "\n" +
                    "  Error = \n" + e.Error.ToString(),
                    "Error processing record",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button2);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    e.Cancel = true;
                    ssStatusText.Text = "Error processing record";
                    CancelMerge();
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
                                "Please specify the Microsoft Excel workbook to use as a datasource.", "Could not open datasource", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnDatasourceBrowse_Click(this, EventArgs.Empty);
            }
            else
            {
                this.names = (string[])e.Data;
                clbNames.Items.Clear();
                foreach (string name in names)
                {
                    clbNames.Items.Add(name);
                    if (this.initialnames == null || this.initialnames.Contains(name))
                    {
                        clbNames.SetItemCheckState(clbNames.Items.IndexOf(name), CheckState.Checked);
                    }
                }
                this.initialnames = null;

                datasourcegood = true;
                btnMerge.Enabled = true;
                btnPrint.Enabled = true;
                ssStatusText.Text = "Ready";
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
                                "Please specify the Microsoft Publisher template to use.", "Could not open template", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        Microsoft.Office.Interop.Publisher.PageSetup setup = (Microsoft.Office.Interop.Publisher.PageSetup)e.Data;
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
                AutoFillDatasource();
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

    private void ReportCardWindowsGUIMerger_FormClosing(object sender, FormClosingEventArgs e)
    {
        worker.Quit();
    }

    private void cbSelectAll_CheckedChanged(object sender, EventArgs e)
    {
        if (cbSelectAll.CheckState != CheckState.Indeterminate)
        {
            CheckState state = cbSelectAll.CheckState;
            for (int i = 0; i < clbNames.Items.Count; i++){
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

