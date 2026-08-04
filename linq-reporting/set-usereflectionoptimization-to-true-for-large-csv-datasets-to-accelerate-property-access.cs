using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // Provides ReportingEngine and CsvDataSource
using Aspose.Words.Reporting; // For CsvDataLoadOptions

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        string[] csvLines =
        {
            "Id,Name,Value",
            "1,Alpha,100",
            "2,Beta,200",
            "3,Gamma,300"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create a template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Enable reflection optimization for large CSV datasets.
        ReportingEngine.UseReflectionOptimization = true;

        // Configure CSV loading options – the first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true; // Ensure column names are recognized.

        // Load CSV data as a data source using the configured options.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the data source named "data".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "data");

        // Save the generated report.
        doc.Save("ReportFromCsv.docx");
    }
}
