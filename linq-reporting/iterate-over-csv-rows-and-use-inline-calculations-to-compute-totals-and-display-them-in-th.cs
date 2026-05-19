using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Quantity,Price",
            "Apple,3,0.5",
            "Banana,2,0.3",
            "Orange,5,0.4"
        });

        // Configure CSV loading options (first line contains headers).
        var loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true;

        // Create CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("-------------------------------------------------");
        // Iterate over CSV rows.
        builder.Writeln("<<foreach [row in items]>>");
        builder.Writeln("Item: <<[row.Name]>> | Qty: <<[row.Quantity]>> | Price: <<[row.Price]>> | Line Total: <<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");
        // Grand total calculated inline using LINQ Sum.
        builder.Writeln("Grand Total: <<[items.Sum(r => r.Quantity * r.Price)]>>");

        // Build the report using the CSV data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "items");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
