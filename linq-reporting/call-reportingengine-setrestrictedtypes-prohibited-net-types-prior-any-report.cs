using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a simple template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Template tries to access System.Type members – this will be blocked.
            builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

            // Restrict access to the System.Type type (and its derived members) before any report is built.
            // This must be done once, typically at application startup.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw when it encounters the blocked type.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report. The restricted type prevents access to GetType().
            engine.BuildReport(doc, new object());

            // The resulting document will contain an empty string where the restricted type was used.
            Console.WriteLine("Resulting document text:");
            Console.WriteLine(doc.GetText().Trim());
        }
    }
}
