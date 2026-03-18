using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a simple template document that attempts to access System.Type members.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The template tries to call GetType() which we will block.
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

        // Restrict the System.Type type so its members cannot be accessed from the template.
        // This must be done before the first report build.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Configure the reporting engine to ignore missing members (the blocked call will yield an empty string).
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report with an empty data source.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("RestrictedReport.docx");
    }
}
