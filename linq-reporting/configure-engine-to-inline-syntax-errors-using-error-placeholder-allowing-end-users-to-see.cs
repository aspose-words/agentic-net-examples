using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Name = "John Doe",
            Age = 30
        };

        // Serialize model to JSON (demonstrates Newtonsoft.Json usage).
        string json = JsonConvert.SerializeObject(model);
        Console.WriteLine($"Model JSON: {json}");

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Correct tag.
        builder.Writeln("Customer Name: <<[model.Name]>>");
        // Faulty tag – references a non‑existent property to trigger an inline error.
        builder.Writeln("Faulty Tag: <<[model.NonExisting]>>");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report_Output.docx";
        templateDoc.Save(outputPath);

        // Output result.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Data model used by the template.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
