using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioTestAppOfficeInteropNative
{
    /// <summary>
    /// This code was mostly taken directly from https://saveenr.gitbooks.io/visioautomation_docs/visio-and-c.html
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;
            var shape = page.DrawRectangle(1, 1, 5, 4);
            shape.Text = "Office Interop Native";
        }
    }
}
