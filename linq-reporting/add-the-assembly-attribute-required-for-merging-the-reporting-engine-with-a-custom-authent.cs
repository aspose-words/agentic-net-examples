using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class AuthHelper
{
    // Example static method that could represent a custom authentication lookup.
    public static string GetUserName()
    {
        return "JohnDoe";
    }
}

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that calls the static method defined above.
        builder.Writeln("Authenticated user: <<[AuthHelper.GetUserName()]>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Register the custom type with the reporting engine (instead of using the missing attribute).
        engine.KnownTypes.Add(typeof(AuthHelper));

        // No data source is required because the template only uses a static call.
        // An empty anonymous object is supplied as the root data source.
        engine.BuildReport(report, new object(), "");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        report.Save(reportPath);
    }
}
