using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class TemplateSecurityExample
{
    static void Main()
    {
        // Create a simple document that will be used as a template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[System.IO.File.ReadAllText(\"secret.txt\")]>>"); // Example of a potentially unsafe call.
        builder.Writeln("Report generated at: <<[DateTime.Now]>>");

        // Restrict types that provide file‑system write capabilities.
        // These types will be inaccessible from the reporting engine's template syntax.
        ReportingEngine.SetRestrictedTypes(
            typeof(System.IO.File),
            typeof(System.IO.FileInfo),
            typeof(System.IO.StreamWriter),
            typeof(System.IO.StreamReader),
            typeof(System.IO.Directory));

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so that attempts to use restricted members do not throw,
            // but instead render as empty strings (or you could omit this option to raise errors).
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The template will be processed, but any access to the restricted
        // types (e.g., System.IO.File) will be blocked.
        engine.BuildReport(template, new { });

        // Save the resulting document.
        template.Save("SecureReport.docx");
    }
}
