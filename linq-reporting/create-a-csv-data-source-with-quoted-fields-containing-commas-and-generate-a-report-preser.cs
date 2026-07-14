using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvReport
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Directory.GetCurrentDirectory();

            // 1. Create a CSV file with quoted fields that contain commas.
            string csvPath = Path.Combine(workDir, "sample.csv");
            File.WriteAllText(csvPath,
                "Id,Description\r\n" +
                "1,\"Item, with, commas\"\r\n" +
                "2,\"Another, item\"\r\n");

            // 2. Build a Word template that uses LINQ Reporting tags.
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("CSV Report");
            builder.Writeln("<<foreach [row in data]>>");
            builder.Writeln("Id: <<[row.Id]>>, Description: <<[row.Description]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 4. Configure CSV loading options (headers present, quote character is double quote).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                QuoteChar = '"',   // Preserve quoted commas.
                Delimiter = ',',   // Default delimiter.
                HasHeaders = true
            };

            // 5. Create the CSV data source.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // 6. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed.
            engine.BuildReport(reportDoc, csvDataSource, "data");

            // 7. Save the generated report.
            string outputPath = Path.Combine(workDir, "report.docx");
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine("Report generated: " + outputPath);
        }
    }
}
