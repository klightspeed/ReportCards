using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SouthernCluster.ReportCards
{
    public static class AppMain
    {
        [STAThread]
        private static int Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolve);

            // Determine if we need to re-exec as 32-bit
            if (!CanUseExcelOdbc())
            {
                if (IntPtr.Size == 8)
                {
                    NativeMethods.TryAttachConsole(-1);
                    Stream run32strm = Assembly.GetExecutingAssembly().GetManifestResourceStream(typeof(AppMain).Namespace + ".Dependencies.Run32.exe");
                    byte[] run32data = new byte[run32strm.Length];
                    run32strm.Read(run32data, 0, (int)run32strm.Length);
                    string run32path = Path.GetTempFileName() + ".exe";
                    // TODO: copy 32-bit wrapper to temp directory
                    File.WriteAllBytes(run32path, run32data);
                    // TODO: execute 32-bit wrapper
                    Process run32proc = Process.Start(run32path, Environment.CommandLine);
                    run32proc.WaitForExit();
                    return run32proc.ExitCode;
                }
                else
                {
                    MessageBox.Show("Unable to use the Microsoft Excel ODBC driver");
                    return 1;
                }
            }

            try
            {
                int ret = ReportCardWindowsGUIMerger.Main(args);
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

        static bool CanUseExcelOdbc()
        {
            try
            {
                string connstring = @"Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=NUL:";
                using (OdbcConnection conn = new OdbcConnection(connstring))
                {
                    conn.Open();
                }
                return false;
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith("ERROR [IM002]"))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                AssemblyName asmname = new AssemblyName(args.Name);
                string resourceName = typeof(AppMain) + ".Dependencies." + asmname.Name + ".dll.gz";
                using (Stream asmgzstream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    using (GZipStream asmstream = new GZipStream(asmgzstream, CompressionMode.Decompress))
                    {
                        byte[] asmdata = new byte[asmstream.Length];
                        asmstream.Read(asmdata, 0, (int)asmstream.Length);
                        return Assembly.Load(asmdata);
                    }
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
