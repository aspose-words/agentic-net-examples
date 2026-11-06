using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingCsvSumExample
{
    public static void Main()
    {
        // Register code page provider for CSV handling (required for non‑UTF8 encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample CSV data.
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        File.WriteAllText(csvPath, "Value1,Value2\n10,20\n5,7\n12,8");

        // -----------------------------------------------------------------
        // Build the template document containing LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV rows (named "row").
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Value1: <<[row.Value1]>>");
        builder.Writeln("Value2: <<[row.Value2]>>");
        // Calculated field – sum of the two columns.
        builder.Writeln("Sum: <<[row.Value1 + row.Value2]>>");
        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options. The default separator is a comma,
        // so we only need to specify that the first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Create the CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, csvData, "data");

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
