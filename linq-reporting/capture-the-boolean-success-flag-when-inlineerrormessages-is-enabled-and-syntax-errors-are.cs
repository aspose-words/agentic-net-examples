using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data model.
        var model = new SampleModel
        {
            Name = "Alice",
            Items = new List<string> { "Item1", "Item2" }
        };

        // Create a template document with a deliberate syntax error (missing closing foreach tag).
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, model, "model");

        // Output the success flag.
        Console.WriteLine($"BuildReport success: {success}");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(reportPath);
    }

    // Creates a template document with a syntax error for demonstration.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Correct tag.
        builder.Writeln("Hello, <<[model.Name]>>!");

        // Intentional syntax error: foreach tag without a closing tag.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("Item: <<[item]>>");
        // Add the missing closing tag to make the template valid.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Sample data model used by the template.
    public class SampleModel
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }
}
