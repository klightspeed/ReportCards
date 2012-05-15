using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Run32
{
    class Program
    {
        static int Main(string[] args)
        {
            try
            {
                Assembly asm = Assembly.LoadFile(args[0]);
                object ret = asm.EntryPoint.Invoke(null, args.Skip(1).ToArray());
                if (ret is int)
                {
                    return (int)ret;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Caught exception executing\n{0}\n{1}", Environment.CommandLine, ex.ToString()), "Caught exception");
                return 1;
            }
        }
    }
}
