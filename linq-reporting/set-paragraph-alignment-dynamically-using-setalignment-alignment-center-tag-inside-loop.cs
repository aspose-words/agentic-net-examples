using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "First Item" },
                new Item { Name = "Second Item" },
                new Item { Name = "Third Item" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a heading.
        builder.Writeln("Dynamic Alignment Report");
        builder.Writeln();

        // Start a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Each item will be placed in its own paragraph.
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // After the report is built, set paragraph alignment dynamically.
        // Here we center all paragraphs that contain item names.
        foreach (Paragraph para in reportDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            string text = para.GetText().Trim();
            if (model.Items.Exists(i => i.Name == text))
            {
                para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            }
        }

        // Save the final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
}
