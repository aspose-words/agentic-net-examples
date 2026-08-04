using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare data model with an empty value.
        var model = new ReportModel
        {
            EmptyValue = string.Empty,
            Name = "John Doe"
        };

        // Configure reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);

        // Verify that no empty paragraphs remain.
        bool hasEmptyParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Any(p => string.IsNullOrWhiteSpace(p.GetText()));

        Console.WriteLine(hasEmptyParagraphs
            ? "Test failed: Empty paragraphs were not removed."
            : "Test passed: Empty paragraphs were successfully removed.");
    }

    private static void CreateTemplate(string filePath)
    {
        // Build a simple template with a tag that will be empty after processing.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Paragraph that will contain an empty value.
        builder.Writeln("<<[model.EmptyValue]>>");

        // Paragraph with actual content to ensure document is not empty.
        builder.Writeln("Hello <<[model.Name]>>!");

        template.Save(filePath);
    }

    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        public string EmptyValue { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }
}
