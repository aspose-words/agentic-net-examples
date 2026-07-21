using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath,
            "Value1,Value2,Description\r\n" +
            "10,20,First row\r\n" +
            "5,7,Second row\r\n" +
            "12,8,Third row");

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("LINQ Reporting – CSV Example");
        builder.Writeln("-------------------------------------------------");

        // Begin a foreach loop over the CSV data source named "csv".
        builder.Writeln("<<foreach [row in csv]>>");
        builder.Writeln("Value 1: <<[row.Value1]>>");
        builder.Writeln("Value 2: <<[row.Value2]>>");
        // Calculated field that sums the two numeric columns.
        builder.Writeln("Sum (Value1 + Value2): <<[row.Value1 + row.Value2]>>");
        builder.Writeln("Description: <<[row.Description]>>");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure CSV loading options – the file has a header row.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // Pass the data source name "csv" to match the foreach tag.
        engine.BuildReport(doc, csvDataSource, "csv");

        // Save the generated report.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportResult.docx");
        doc.Save(resultPath);
    }
}
