using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBatch
{
    public class Program
    {
        public static void Main()
        {
            // Working directory.
            string workDir = Directory.GetCurrentDirectory();

            // 1. Create sample CSV data.
            string csvPath = Path.Combine(workDir, "customers.csv");
            CreateSampleCsv(csvPath);

            // 2. Create a LINQ Reporting template programmatically.
            string templatePath = Path.Combine(workDir, "template.docx");
            CreateTemplateDocument(templatePath);

            // 3. Load the template.
            Document templateDoc = new Document(templatePath);

            // 4. Prepare CSV data source with headers.
            var loadOptions = new CsvDataLoadOptions(true);
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // 5. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            // The root data source name must match the name used in the template tags ("customers").
            engine.BuildReport(templateDoc, csvDataSource, "customers");

            // 6. Save the generated report.
            string outputPath = Path.Combine(workDir, "CustomerReports.docx");
            templateDoc.Save(outputPath);
        }

        // Creates a simple CSV file with a few customer records.
        private static void CreateSampleCsv(string path)
        {
            string[] lines =
            {
                "CustomerName,Address,Email",
                "Alice Johnson,123 Maple St.,alice@example.com",
                "Bob Smith,456 Oak Ave.,bob@example.com",
                "Carol Davis,789 Pine Rd.,carol@example.com"
            };
            File.WriteAllLines(path, lines);
        }

        // Builds a Word template that uses LINQ Reporting tags to iterate over customers.
        private static void CreateTemplateDocument(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title
            builder.Writeln("Customer Report");
            builder.Writeln("----------------");

            // Begin foreach loop over the CSV rows (exposed as "customers").
            builder.Writeln("<<foreach [c in customers]>>");

            // Individual customer fields.
            builder.Writeln("Name   : <<[c.CustomerName]>>");
            builder.Writeln("Address: <<[c.Address]>>");
            builder.Writeln("Email  : <<[c.Email]>>");
            builder.Writeln(""); // Blank line between records.

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(path);
        }
    }
}
