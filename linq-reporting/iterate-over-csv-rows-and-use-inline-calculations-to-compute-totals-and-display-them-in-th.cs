using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Item,Quantity,UnitPrice\r\n" +
            "Apple,3,0.50\r\n" +
            "Banana,5,0.30\r\n" +
            "Orange,2,0.80\r\n");

        // Load CSV rows into a strongly‑typed list.
        var model = new ReportModel();
        foreach (var line in File.ReadAllLines(csvPath))
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("Item")) continue; // Skip header.
            var parts = line.Split(',');
            if (parts.Length != 3) continue;

            var row = new CsvRow
            {
                Item = parts[0],
                Quantity = int.Parse(parts[1], CultureInfo.InvariantCulture),
                UnitPrice = decimal.Parse(parts[2], CultureInfo.InvariantCulture)
            };
            model.Items.Add(row);
            model.GrandTotal += row.Quantity * row.UnitPrice;
        }

        // Create the LINQ Reporting template.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Sales Report");
        builder.Writeln();

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Unit Price");
        builder.InsertCell(); builder.Writeln("Line Total");
        builder.EndRow();

        // Data rows.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Item]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
        builder.EndRow();
        builder.Writeln("<</foreach>>");

        // Grand total row.
        builder.InsertCell(); builder.Writeln("Grand Total");
        builder.InsertCell(); builder.Writeln("");
        builder.InsertCell(); builder.Writeln("");
        builder.InsertCell(); builder.Writeln("<<[model.GrandTotal]>>");
        builder.EndRow();

        builder.EndTable();

        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<CsvRow> Items { get; set; } = new();
    public decimal GrandTotal { get; set; }
}

// CSV row representation.
public class CsvRow
{
    public string Item { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
