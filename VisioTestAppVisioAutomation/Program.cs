using ConnectorType = VisioAutomation.Models.ConnectorType;

namespace VisioTestAppVisioAutomation
{
    /// <summary>
    /// https://stackoverflow.com/questions/2258882/generate-visio-diagram-on-the-fly-with-net
    /// The accepted code snippet in the SO link above is not valid C#, and it also appears to reference out of date code. I had to adjust to get it to work.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            var visapp = new Microsoft.Office.Interop.Visio.Application();
            var d = new VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphLayout();
            var basic_stencil = "basic_u.vss";
            var n0 = d.AddShape("n0", "Node 0", basic_stencil, "Rectangle");
            var n1 = d.AddShape("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = d.AddShape("n2", "Node 2", basic_stencil, "Rectangle");
            var n3 = d.AddShape("n3", "Node 3", basic_stencil, "Rectangle");
            var n4 = d.AddShape("n4", "Node 4\nUnconnected", basic_stencil, "Rectangle");
            var c0 = d.AddConnection("c0", n0, n1, "0 -> 1", ConnectorType.Curved);
            var c1 = d.AddConnection("c1", n1, n2, "1 -> 2", ConnectorType.RightAngle);
            var c2 = d.AddConnection("c2", n1, n0, "0 -> 1", ConnectorType.Curved);
            var c3 = d.AddConnection("c3", n0, n2, "0 -> 2", ConnectorType.Straight);
            var c4 = d.AddConnection("c4", n2, n3, "2 -> 3", ConnectorType.Curved);
            var c5 = d.AddConnection("c5", n3, n0, "3 -> 0", ConnectorType.Curved);
            var options = new VisioAutomation.Models.Layouts.DirectedGraph.MsaglLayoutOptions();
            visapp.Documents.Add("");
            var page = visapp.ActivePage;
            d.Render(page, options);
        }
    }
}
