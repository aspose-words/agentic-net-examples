using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // ReportingEngine, CsvDataSource, CsvDataLoadOptions

namespace AsposeWordsLinqReportingCsvExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV parsing (required for non‑UTF8 encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // File paths for the sample CSV, template, and output report.
            string csvPath = "sample.csv";
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // Create a simple CSV file with a header row and three data rows.
            File.WriteAllLines(csvPath, new[]
            {
                "Id,Name,Age",
                "1,John Doe,30",
                "2,Jane Smith,25",
                "3,Bob Johnson,40"
            }, Encoding.UTF8);

            // Build a Word template that contains LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // LINQ Reporting foreach loop over the CSV data source named "data".
            builder.Writeln("<<foreach [row in data]>>");
            builder.Writeln("Id: <<[row.Id]>>");
            builder.Writeln("Name: <<[row.Name]>>");
            builder.Writeln("Age: <<[row.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template document that will be populated.
            Document reportDoc = new Document(templatePath);

            // Configure CSV loading options – the first line contains column headers.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);

            // Create a CSV data source from the file using the configured options.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(reportDoc, csvDataSource, "data");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
