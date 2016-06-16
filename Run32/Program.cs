using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;

namespace Run32
{
    class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            try
            {
                string program = Path.GetFullPath(args[0]);

                if (program.EndsWith(".vshost.exe"))
                {
                    program = program.Substring(0, program.Length - ".vshost.exe".Length) + ".exe";
                }

                Assembly asm = Assembly.LoadFile(program);
                MethodInfo entrypoint = asm.EntryPoint;
                /*
                MessageBox.Show(
                    String.Format("{0} {1}.{2}({3});",
                        entrypoint.ReturnType.ToString(),
                        entrypoint.ReflectedType.ToString(),
                        entrypoint.Name,
                        String.Join(", ",
                            entrypoint.GetParameters().Select(p => 
                                (p.IsIn ? (p.IsOut ? "ref" : "") : (p.IsOut ? "out " : "")) +
                                p.ParameterType.ToString() + " " +
                                p.Name
                            ).ToArray()
                        )
                    )
                );
                 */
                object ret = entrypoint.Invoke(null, new object[] { args.Skip(1).ToArray() });
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
