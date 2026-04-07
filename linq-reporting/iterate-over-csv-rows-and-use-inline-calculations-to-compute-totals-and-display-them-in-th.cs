using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple wrapper to hold the grand total that will be displayed after the table.
    public class ReportInfo
    {
        public double GrandTotal { get; set; } = 0.0;
    }

    public static void Main()
    {
        // Register code page provider for CSV parsing on .NET Core/.NET 5+.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers: Item,Quantity,Price
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        string[] csvLines =
        {
            "Item,Quantity,Price",
            "Apple,10,0.5",
            "Banana,5,0.3",
            "Orange,8,0.6"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Load the CSV as a data source for the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Compute the grand total (Quantity * Price) using LINQ.
        // -----------------------------------------------------------------
        double grandTotal = csvLines
            .Skip(1) // skip header
            .Select(line => line.Split(','))
            .Select(parts => new
            {
                Quantity = double.Parse(parts[1]),
                Price = double.Parse(parts[2])
            })
            .Sum(item => item.Quantity * item.Price);

        var reportInfo = new ReportInfo { GrandTotal = grandTotal };

        // -----------------------------------------------------------------
        // 4. Build the template document programmatically.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Sales Report");
        builder.Writeln("Item\tQty\tPrice\tTotal");
        builder.Writeln("<<foreach [row in data]>>");
        // Inline calculation for line total: Quantity * Price
        builder.Writeln("<<[row.Item]>>\t<<[row.Quantity]>>\t<<[row.Price]>>\t<<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Grand Total: <<[info.GrandTotal]>>");

        // -----------------------------------------------------------------
        // 5. Run the reporting engine with both data sources.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(
            doc,
            new object[] { csvDataSource, reportInfo },
            new string[] { "data", "info" });

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SalesReport.docx");
        doc.Save(outputPath);
    }
}
