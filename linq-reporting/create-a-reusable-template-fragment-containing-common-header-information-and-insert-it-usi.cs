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
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string templateFolder = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
        Directory.CreateDirectory(templateFolder);

        // Create header fragment file
        string headerPath = Path.Combine(templateFolder, "header.txt");
        File.WriteAllText(headerPath, "=== Report Header ===\nTitle: <<[model.Title]>>\n");

        // Create main template document
        string templatePath = Path.Combine(templateFolder, "report_template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert header fragment content
        string headerContent = File.ReadAllText(headerPath);
        builder.Writeln(headerContent);

        // Add a blank line
        builder.Writeln();

        // Insert a foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new()
        {
            Title = "Sample Report",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report_output.docx");
        reportDoc.Save(outputPath);
    }
}

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
