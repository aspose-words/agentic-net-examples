using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // 1. Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated with Aspose.Words LINQ Reporting Engine");
        builder.Writeln();

        // Use a supported <<doc>> tag to include an external document.
        // The IncludeSource.Document property will be null because the file does not exist,
        // and the engine will skip it when AllowMissingMembers is enabled.
        builder.Writeln("<<doc [src.Document]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare the data source.
        var src = new IncludeSource
        {
            // Attempt to load a non‑existent file; result will be null.
            Document = File.Exists("nonexistent.docx") ? new Document("nonexistent.docx") : null
        };

        // 4. Configure the ReportingEngine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty
        };

        // Build the report using the wrapper object as the data source.
        engine.BuildReport(doc, src, "src");

        // 5. Save the generated report.
        doc.Save(reportPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
    }

    // Wrapper class exposing a Document property for the <<doc>> tag.
    public class IncludeSource
    {
        public Document? Document { get; set; }
    }
}
