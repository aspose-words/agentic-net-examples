using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Load the template document.
        Document templateDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            Name = "John Doe",
            Items = new() { "Item A", "Item B", "Item C" }
        };

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "output.docx");
        templateDoc.Save(resultPath);

        // Indicate the result.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Output saved to: {resultPath}");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Simple property tag (valid).
        builder.Writeln("Customer Name: <<[model.Name]>>");

        // Invalid property tag (will trigger an inline error).
        builder.Writeln("Missing Property: <<[model.Unknown]>>");

        // Loop over a collection.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }

    // Data model used by the report.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }
}
