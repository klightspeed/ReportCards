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
        public static string EscapeCommandLineArgument(string arg)
        {
            StringBuilder sb = new StringBuilder();
            StringReader rdr = new StringReader(arg);
            int c;
            sb.Append('"');

            while ((c = rdr.Read()) > 0)
            {
                if (c == '"')
                {
                    sb.Append("\\\"");
                }
                else if (c == '\\')
                {
                    int nrbackslash = 1;

                    while (rdr.Peek() == '\\')
                    {
                        nrbackslash++;
                        rdr.Read();
                    }

                    if (rdr.Peek() == '"')
                    {
                        sb.Append(new String('\\', nrbackslash * 2));
                    }
                    else
                    {
                        sb.Append(new String('\\', nrbackslash));
                    }
                }
                else
                {
                    sb.Append((char)c);
                }
            }

            sb.Append('"');
            return sb.ToString();
        }

        public static string CreateCommandArguments(params string[] args)
        {
            return String.Join(" ", args.Select(s => EscapeCommandLineArgument(s)).ToArray());
        }

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

#if false
                    File.Copy(Assembly.GetExecutingAssembly().Location, run32path);
                    string corflagspath = Path.GetTempFileName() + ".exe";

                    using (Stream corflagsstrm = Assembly.GetExecutingAssembly().GetManifestResourceStream(typeof(AppMain).Namespace + ".Dependencies.CorFlags.exe"))
                    {
                        byte[] corflagsdata = new byte[corflagsstrm.Length];
                        corflagsstrm.Read(corflagsdata, 0, (int)corflagsstrm.Length);
                        File.WriteAllBytes(corflagspath, corflagsdata);
                    }

                    Process corflagsproc = Process.Start(corflagspath, CreateCommandArguments("/32BIT+", run32path));
                    corflagsproc.WaitForExit();
                    Process run32proc = Process.Start(run32path, CreateCommandArguments(args));
                    run32proc.WaitForExit();
                    return run32proc.ExitCode;
#else
                    string run32path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Run32.exe");
                    using (Stream run32strm = Assembly.GetExecutingAssembly().GetManifestResourceStream(typeof(AppMain).Namespace + ".Dependencies.Run32.exe"))
                    {
                        byte[] run32data = new byte[run32strm.Length];
                        run32strm.Read(run32data, 0, (int)run32strm.Length);

                        try
                        {
                            File.WriteAllBytes(run32path, run32data);
                        }
                        catch
                        {
                            run32path = Path.GetTempFileName() + ".exe";
                            File.WriteAllBytes(run32path, run32data);
                        }
                    }

                    Process run32proc = Process.Start(run32path, CreateCommandArguments(new string[] { Assembly.GetExecutingAssembly().Location }.Concat(args).ToArray()));
                    run32proc.WaitForExit();
                    return run32proc.ExitCode;
#endif
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
