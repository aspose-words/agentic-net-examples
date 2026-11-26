using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Items Report");
        builder.Writeln();

        // Begin foreach loop over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Item Name");
        builder.EndRow();

        // Data row with dynamic background color.
        builder.InsertCell();
        builder.Writeln("<<backColor [item.Color]>> <<[item.Name]>> <</backColor>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // Sample data model.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Apple",  Color = "LightGreen" },
                new Item { Name = "Banana", Color = "#FFFF00" },
                new Item { Name = "Cherry", Color = "Red" }
            }
        };

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
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
    public string Color { get; set; } = "";
}
