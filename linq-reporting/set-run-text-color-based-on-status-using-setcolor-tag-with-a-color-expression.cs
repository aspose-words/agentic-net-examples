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
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template document.
        Document templateDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Status = "Good", Text = "Operation succeeded." },
                new Item { Status = "Bad", Text = "Operation failed." },
                new Item { Status = "Pending", Text = "Operation pending." }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        templateDoc.Save(reportPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple paragraph with a foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(
            "<<textColor [item.Status == \"Good\" ? \"Green\" : item.Status == \"Bad\" ? \"Red\" : \"Black\"]>>" +
            "<<[item.Text]>>" +
            "<</textColor>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Status { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
