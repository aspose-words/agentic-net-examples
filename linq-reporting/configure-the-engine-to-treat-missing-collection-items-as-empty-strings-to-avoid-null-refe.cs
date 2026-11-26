using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare sample data with missing items replaced by empty strings.
        ReportModel model = new()
        {
            Items = new List<string> { "First", string.Empty, "Third", string.Empty, "Fifth" }
        };

        // Configure the reporting engine (no special options needed for empty strings).
        ReportingEngine engine = new();

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to: {outputPath}");
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Items List:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - <<[item]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    public class ReportModel
    {
        public List<string> Items { get; set; } = new();
    }
}
