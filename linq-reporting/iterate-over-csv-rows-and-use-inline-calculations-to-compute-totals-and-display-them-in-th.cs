using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Quantity,Price",
            "Apple,3,0.5",
            "Banana,2,0.3",
            "Orange,5,0.4"
        });

        // Create a simple template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Item: <<[row.Item]>>, Qty: <<[row.Quantity]>>, Price: <<[row.Price]>>, Total: <<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("Grand Total: <<[summary.Total]>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Compute grand total by reading the CSV file.
        double grandTotal = 0;
        string[] lines = File.ReadAllLines(csvPath);
        for (int i = 1; i < lines.Length; i++) // Skip header.
        {
            string[] parts = lines[i].Split(',');
            if (parts.Length >= 3 &&
                double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double qty) &&
                double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out double price))
            {
                grandTotal += qty * price;
            }
        }

        // Summary object to expose the total.
        Summary summary = new Summary { Total = grandTotal };

        // Build the report using two data sources: "data" (CSV) and "summary".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, new object[] { csvData, summary }, new[] { "data", "summary" });

        // Save the final report.
        reportDoc.Save("report.docx");
    }
}

// Simple wrapper class for the grand total.
public class Summary
{
    public double Total { get; set; } = 0;
}
