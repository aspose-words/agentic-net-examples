using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string csvPath = Path.Combine(outputDir, "data.csv");
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // 1. Create sample CSV data.
        File.WriteAllLines(csvPath, new[]
        {
            "Value1,Value2",
            "10,20",
            "5,7",
            "12,8"
        });

        // 2. Create the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV rows (named "data").
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Value1: <<[row.Value1]>>");
        builder.Writeln("Value2: <<[row.Value2]>>");
        // Calculated field: sum of the two numeric columns.
        builder.Writeln("Sum: <<[row.Value1 + row.Value2]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Prepare CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            Delimiter = ','
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // The data source name "data" must match the name used in the template tags.
        engine.BuildReport(reportDoc, csvData, "data");

        // 6. Save the generated report.
        reportDoc.Save(reportPath);
    }
}
