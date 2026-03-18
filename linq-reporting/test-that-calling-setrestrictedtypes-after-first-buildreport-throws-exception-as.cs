using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Simple template that tries to access a restricted type.
        string template = "<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>";
        Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(template)));

        // Set the restricted type before the first report build (allowed).
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Build the report with the AllowMissingMembers option to avoid template errors.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.BuildReport(doc, new object());

        // After the first BuildReport call, attempting to modify restricted types must throw.
        try
        {
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));
            Console.WriteLine("FAIL: No exception was thrown.");
        }
        catch (ArgumentException)
        {
            Console.WriteLine("PASS: ArgumentException thrown as documented.");
        }
    }
}
