using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string csvPath = Path.Combine(outputDir, "data.csv");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV rows (data source name: data).
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Value1: <<[row.Value1]>>");
        builder.Writeln("Value2: <<[row.Value2]>>");
        // Calculated field: sum of the two numeric columns.
        builder.Writeln("Sum: <<[row.Value1 + row.Value2]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a sample CSV file with two numeric columns.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Value1,Value2",
            "10,20",
            "5,7",
            "12,8"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading to treat the first line as column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report. The data source name used in the template is "data".
        engine.BuildReport(doc, csvData, "data");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);

        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
