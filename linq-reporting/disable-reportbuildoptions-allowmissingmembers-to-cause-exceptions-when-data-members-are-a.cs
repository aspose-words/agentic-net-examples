using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Existing property; the template will reference a non‑existent one.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create a template document with a tag that refers to a missing member.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Customer: <<[model.Name]>>");
        builder.Writeln("Missing: <<[model.MissingProperty]>>"); // This member does not exist.

        // Save and reload the template to simulate a real file scenario.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // Prepare the data model (does not contain MissingProperty).
        var model = new Model { Name = "John Doe" };

        // Configure the reporting engine without AllowMissingMembers.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Disallow missing members.

        try
        {
            // Attempt to build the report; should throw because of the missing member.
            engine.BuildReport(loadedTemplate, model, "model");
            // If no exception, save the generated report.
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Expected path: report generation fails due to missing member.
            Console.WriteLine($"Report generation failed: {ex.Message}");
        }
    }
}
