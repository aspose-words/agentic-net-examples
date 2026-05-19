using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a sample CSV file with headers and numeric values.
        string csvPath = "sample.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Item,Value1,Value2",
            "Apple,10,5",
            "Banana,7,3",
            "Cherry,12,8"
        });

        // Build a template document containing LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Item: <<[row.Item]>>, Sum: <<[row.Value1 + row.Value2]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template document for report generation.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source options (headers present, comma delimiter).
        var loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };

        // Create the CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // Save the final report.
        reportDoc.Save("report.docx");
    }
}
