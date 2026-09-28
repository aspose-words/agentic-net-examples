using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class CsvRow
{
    public string Item { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<CsvRow> Rows { get; set; } = new();
    public decimal GrandTotal { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file.
        string csvPath = Path.Combine(outputDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Quantity,Price",
            "Apple,10,0.5",
            "Banana,5,0.3",
            "Orange,8,0.6"
        });

        // Load CSV data into the model.
        ReportModel model = new();
        string[] lines = File.ReadAllLines(csvPath);
        for (int i = 1; i < lines.Length; i++)
        {
            string[] parts = lines[i].Split(',');
            if (parts.Length != 3) continue;

            model.Rows.Add(new CsvRow
            {
                Item = parts[0],
                Quantity = int.Parse(parts[1], CultureInfo.InvariantCulture),
                Price = decimal.Parse(parts[2], CultureInfo.InvariantCulture)
            });
        }

        model.GrandTotal = model.Rows.Sum(r => r.Quantity * r.Price);

        // Build the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Sales Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [row in Rows]>>");
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Data row with inline calculation.
        builder.InsertCell(); builder.Writeln("<<[row.Item]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Quantity * row.Price]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Grand Total: <<[GrandTotal]>>");

        // Save the template and reload it for reporting.
        templateDoc.Save(templatePath);
        Document doc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);
    }
}
