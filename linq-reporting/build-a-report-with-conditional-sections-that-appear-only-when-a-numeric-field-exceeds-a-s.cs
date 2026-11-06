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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare template document.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportData model = new()
        {
            Items = new()
            {
                new Item { Name = "Product A", Sales = 1500 },
                new Item { Name = "Product B", Sales = 800 },
                new Item { Name = "Product C", Sales = 2500 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Sales Report");
        builder.Writeln();

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with header
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Sales");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // Conditional section: show "High" when sales > 1000, otherwise "Normal"
        builder.Writeln("<<if [item.Sales > 1000]>>");
        builder.Writeln("<<[item.Sales]>> (High)");
        builder.Writeln("<</if>>");
        builder.Writeln("<<if [item.Sales <= 1000]>>");
        builder.Writeln("<<[item.Sales]>> (Normal)");
        builder.Writeln("<</if>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(path);
    }
}

public class ReportData
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public double Sales { get; set; }
}
