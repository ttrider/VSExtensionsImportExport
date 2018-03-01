using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TTRider.VSExtensionsImportExport
{
    static class Utilities
    {
        public static void WriteLine(this IVsOutputWindowPane pane, string format, params object[] args)
        {
            pane.OutputStringThreadSafe(string.Format(format, args)+"\r\n");
        }

        public static void Write(this IVsOutputWindowPane pane, string format, params object[] args)
        {
            pane.OutputStringThreadSafe(string.Format(format, args));
        }
    }
}
