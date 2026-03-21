using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TemplateSecurityExample
{
    public static void Main()
    {
        // Create a simple document that will be used as a template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[System.IO.File.ReadAllText(\"secret.txt\")]>>"); // Example of a potentially unsafe call.
        builder.Writeln("Report generated at: <<[DateTime.Now]>>");

        // Restrict types that provide file‑system write capabilities.
        ReportingEngine.SetRestrictedTypes(
            typeof(System.IO.File),
            typeof(System.IO.FileInfo),
            typeof(System.IO.StreamWriter),
            typeof(System.IO.StreamReader),
            typeof(System.IO.Directory));

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report with a visible data source (empty object).
        engine.BuildReport(template, new object());

        // Save the resulting document.
        template.Save("SecureReport.docx");
    }
}
