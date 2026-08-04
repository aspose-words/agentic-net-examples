using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Square root of 16: <<[Math.Sqrt(16)]>>");
        builder.Writeln("JSON representation: <<[JsonConvert.SerializeObject(model)]>>");

        // Save and reload the template to simulate typical workflow.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data model.
        var model = new SampleModel { Name = "Alice" };

        // Configure the reporting engine and register external types.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Math));               // Core .NET type.
        engine.KnownTypes.Add(typeof(JsonConvert));        // Third‑party type from Newtonsoft.Json.

        // Build the report.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);
    }
}

// Sample data model with a non‑nullable property.
public class SampleModel
{
    public string Name { get; set; } = string.Empty;
}
