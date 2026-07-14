using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        string templatePath = "Template.docx";
        string outputPath = "ReportOutput.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Product Report");

        // Header table
        builder.StartTable();

        // First header row with horizontal merge
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>>Product Info");
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>>Product Info");
        builder.InsertCell();
        builder.Write("<<cellMerge>>Quantity");
        builder.EndRow();

        // Sub‑header row
        builder.InsertCell(); builder.Write("Category");
        builder.InsertCell(); builder.Write("Name");
        builder.InsertCell(); builder.Write("Quantity");
        builder.EndRow();

        builder.EndTable();

        // Data rows – placed inside a foreach block
        builder.Writeln("<<foreach [item in Items]>>");
        Table dataTable = builder.StartTable();

        builder.InsertCell(); builder.Write("<<[item.Category]>>");
        builder.InsertCell(); builder.Write("<<[item.Name]>>");
        builder.InsertCell(); builder.Write("<<[item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Prepare sample data
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Category = "Electronics", Name = "Smartphone", Quantity = 15 },
                new Item { Category = "Electronics", Name = "Laptop", Quantity = 7 },
                new Item { Category = "Home", Name = "Coffee Maker", Quantity = 12 },
                new Item { Category = "Home", Name = "Vacuum Cleaner", Quantity = 5 }
            }
        };

        // -----------------------------------------------------------------
        // Build the report
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
