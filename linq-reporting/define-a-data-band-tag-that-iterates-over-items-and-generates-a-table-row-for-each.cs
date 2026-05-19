using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Required for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel();
        model.Items.Add(new Item { Index = 1, Name = "Apple" });
        model.Items.Add(new Item { Index = 2, Name = "Banana" });
        model.Items.Add(new Item { Index = 3, Name = "Cherry" });

        // Build the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a data band that will repeat the rows.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row – this row will be repeated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // Finish the table and the data band.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
