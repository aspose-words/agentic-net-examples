using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title.
        builder.Writeln("Order Report");
        builder.Writeln();

        // Insert a placeholder for the customer name.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln();

        // Build a table that will be repeated for each item in the Items collection.
        builder.Writeln("<<foreach [item in model.Items]>>");
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // 2. Prepare the data model.
        ReportModel model = new ReportModel
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Widget" },
                new Item { Index = 2, Name = "Gadget" },
                new Item { Index = 3, Name = "Doohickey" }
            }
        };

        // 3. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(template, model, "model");

        // 4. Export the rendered document to PDF.
        template.Save("Report.pdf", SaveFormat.Pdf);
    }
}

// Data model classes (public with initialized properties to avoid nullable warnings).
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
