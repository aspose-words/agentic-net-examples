using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvRow
{
    public string Item { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<CsvRow> Rows { get; set; } = new();
    public decimal GrandTotal => Rows.Sum(r => r.Quantity * r.Price);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Quantity,Price",
            "Apple,3,0.5",
            "Banana,5,0.3",
            "Orange,2,0.8"
        });

        // Load CSV rows into the model.
        ReportModel model = new();
        string[] lines = File.ReadAllLines(csvPath);
        for (int i = 1; i < lines.Length; i++) // Skip header.
        {
            string[] parts = lines[i].Split(',');
            if (parts.Length != 3) continue;
            model.Rows.Add(new CsvRow
            {
                Item = parts[0],
                Quantity = int.Parse(parts[1]),
                Price = decimal.Parse(parts[2])
            });
        }

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Item\tQuantity\tPrice\tTotal");
        builder.Writeln("<<foreach [row in Rows]>>");
        builder.Writeln("<<[row.Item]>>\t<<[row.Quantity]>>\t<<[row.Price]>>\t<<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Grand Total: <<[GrandTotal]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
