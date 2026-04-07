using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvReport
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for CSV parsing on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string workDir = Directory.GetCurrentDirectory();
            string csvPath = Path.Combine(workDir, "data.csv");
            string templatePath = Path.Combine(workDir, "template.docx");
            string outputPath = Path.Combine(workDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a CSV file with quoted fields that contain commas.
            // -----------------------------------------------------------------
            // Header line.
            string[] csvLines =
            {
                "Name,Address",
                // Data rows – fields that contain commas are enclosed in double quotes.
                "\"John, Jr.\",\"123 Main St, Apt 4, New York\"",
                "\"Jane Smith\",\"456 Oak Ave, Suite 5, Los Angeles\""
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build a Word template that uses LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("==============");
            // Begin a foreach loop over the CSV rows (exposed as 'persons').
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name   : <<[person.Name]>>");
            builder.Writeln("Address: <<[person.Address]>>");
            builder.Writeln(""); // Empty line between records.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template back (simulating a real‑world scenario).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Configure CSV loading options (quote character, delimiter, headers).
            // -----------------------------------------------------------------
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
            {
                Delimiter = ',',      // default, set explicitly for clarity
                QuoteChar = '"',      // fields are quoted with double quotes
                CommentChar = '#',    // no comment lines in our sample
            };

            // Create the CSV data source.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // -----------------------------------------------------------------
            // 5. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // No special options needed for this simple example.
                Options = ReportBuildOptions.None
            };

            // The data source name used in the template tags is "persons".
            engine.BuildReport(reportDoc, csvDataSource, "persons");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            // The example finishes here; no user interaction is required.
        }
    }
}
