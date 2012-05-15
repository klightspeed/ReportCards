namespace SouthernCluster.ReportCards
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    //using MSXML2;
    using System.Threading;
    using System.Net;

    internal class XmlHttpClient : IDisposable
    {
        //private XMLHTTP req;

        public XmlHttpClient ()
        {
            req = new XMLHTTP();
        }

        ~XmlHttpClient()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (req != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(req);
                }
            }
            req = null;
        }

        public string DownloadString(string url)
        {
            req.open("GET", url, false);
            req.send();
            if (req.status != 200)
            {
                throw new WebException(req.statusText, WebExceptionStatus.ProtocolError);
            }
            return req.responseText;
        }
    }
}
