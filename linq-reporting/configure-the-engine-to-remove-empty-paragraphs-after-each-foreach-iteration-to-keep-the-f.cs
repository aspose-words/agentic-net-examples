using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple" },
                new Item { Name = "Banana" },
                new Item { Name = "Cherry" }
            }
        };

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Items List:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        // An empty paragraph that would normally remain after the foreach.
        builder.Writeln("");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
}
