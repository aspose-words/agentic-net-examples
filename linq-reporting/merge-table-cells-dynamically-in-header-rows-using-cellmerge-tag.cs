using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section
                {
                    HeaderGroup = "Fruits",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple", Quantity = 10 },
                        new Item { Name = "Banana", Quantity = 20 }
                    }
                },
                new Section
                {
                    HeaderGroup = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Name = "Carrot", Quantity = 15 },
                        new Item { Name = "Tomato", Quantity = 25 }
                    }
                }
            }
        };

        // Create template.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Outer foreach – one table per section.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Table for the section header and column titles.
        builder.StartTable();

        // Header row with merged cells.
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[section.HeaderGroup]>>");
        builder.InsertCell();
        builder.Writeln("<<cellMerge>><<[section.HeaderGroup]>>");
        builder.EndRow();

        // Column titles row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        builder.EndTable();

        // Inner foreach – rows for each item.
        builder.Writeln("<<foreach [item in section.Items]>>");
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string HeaderGroup { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
