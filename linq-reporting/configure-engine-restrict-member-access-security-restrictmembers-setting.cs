using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a simple template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The template attempts to access a member of System.Type, which we will restrict.
        builder.Writeln("<<[typeof(System.Type).FullName]>>");

        // Restrict access to System.Type members for the reporting engine.
        // This must be done before the first BuildReport call.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Configure the engine to ignore missing members (so the restricted access does not throw).
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The restricted member will be treated as missing and therefore omitted.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("RestrictedReport.docx");
    }
}
