using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace SouthernCluster.ReportCards
{
    class NativeMethods
    {
        [DllImport("kernel32.dll")]
        internal static extern bool AttachConsole(int pid);

        // Windows.Forms.MessageBox doesn't have Cancel/Try Again/Continue
        [DllImport("user32.dll", EntryPoint = "MessageBoxW", CharSet = CharSet.Unicode)]
        internal static extern int MsgBox(IntPtr hwnd, string text, string caption, int type);

        internal const int MB_CANCELTRYCONTINUE = 0x06;
        internal const int MB_ICONERROR = 0x10;
        internal const int MB_DEFBUTTON2 = 0x100;
        internal const int MB_TASKMODAL = 0x2000;
        internal const int IDCANCEL = 2;
        internal const int IDTRYAGAIN = 10;
        internal const int IDCONTINUE = 11;

        internal static bool TryAttachConsole(int pid)
        {
            try
            {
                return NativeMethods.AttachConsole(pid);
            }
            catch (System.Security.SecurityException)
            {
                return false;
            }
        }
    }
}
