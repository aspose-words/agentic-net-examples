using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
@"Name,Quantity,Price
Apple,3,0.5
Banana,2,0.3
Orange,5,0.4");

        // Build a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Item Report");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("<<foreach [row in items]>>");
        builder.Writeln("Item: <<[row.Name]>>\tQty: <<[row.Quantity]>>\tPrice: <<[row.Price]>>\tTotal: <<[row.Quantity * row.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("Grand Total: <<[items.Sum(r => r.Quantity * r.Price)]>>");

        // Load CSV data source with header row.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, csvData, "items");

        // Save the generated report.
        template.Save("Report.docx");
    }
}
