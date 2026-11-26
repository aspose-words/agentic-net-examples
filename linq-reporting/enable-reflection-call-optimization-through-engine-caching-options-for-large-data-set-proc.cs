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
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "template.docx";
        string reportPath = "report.docx";

        // Create the template document with LINQ Reporting tags.
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare a large data set.
        Order order = new Order();
        for (int i = 1; i <= 1000; i++)
        {
            order.Items.Add(new Item
            {
                Index = i,
                Name = $"Item #{i}"
            });
        }

        // Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(reportPath);

        // Output result status.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to '{reportPath}'.");
    }

    private static void CreateTemplate(string path)
    {
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Large Data Set Report");
        builder.Writeln();

        // Begin foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header and rows.
        Table table = builder.StartTable();
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

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(path);
    }
}

// Data model classes.
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
