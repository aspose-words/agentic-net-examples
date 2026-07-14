using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // 1. Create sample CSV file.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Quantity,Price",
            "Apple,3,0.50",
            "Banana,2,0.30",
            "Orange,5,0.80"
        });

        // 2. Compute grand total from CSV (Quantity * Price).
        decimal grandTotal = 0m;
        foreach (var line in File.ReadAllLines(csvPath).Skip(1))
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;

            if (int.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out int qty) &&
                decimal.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
            {
                grandTotal += qty * price;
            }
        }

        // 3. Create the template document programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Report");
        builder.Writeln();

        // Begin foreach over CSV rows (named "items").
        builder.Writeln("<<foreach [row in items]>>");
        builder.Writeln("Item: <<[row.Item]>>, Qty: <<[row.Quantity]>>, Price: <<[row.Price]>>, Line Total: <<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("Grand Total: <<[summary.Total]>>");

        templateDoc.Save(templatePath);

        // 4. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 5. Prepare data sources.
        var csvLoadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true,
            QuoteChar = '"',
            CommentChar = '#'
        };
        CsvDataSource itemsDataSource = new CsvDataSource(csvPath, csvLoadOptions);

        // Summary object for the grand total.
        Summary summary = new Summary { Total = grandTotal };

        // 6. Build the report using two data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc,
            new object[] { itemsDataSource, summary },
            new string[] { "items", "summary" });

        // 7. Save the final report.
        string outputPath = Path.Combine(workDir, "SalesReport.docx");
        reportDoc.Save(outputPath);
    }
}

// Simple wrapper class for the summary data.
public class Summary
{
    public decimal Total { get; set; }
}
