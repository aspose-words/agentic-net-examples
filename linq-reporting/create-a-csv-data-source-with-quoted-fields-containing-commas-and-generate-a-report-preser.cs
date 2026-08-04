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
            // Register code page provider for CSV encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string workDir = Directory.GetCurrentDirectory();
            string csvPath = Path.Combine(workDir, "data.csv");
            string templatePath = Path.Combine(workDir, "template.docx");
            string reportPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a CSV file with quoted fields that contain commas.
            // -----------------------------------------------------------------
            // Header: Name,Address
            // Data rows: "Doe, John","123 Main St, Apt 4"
            // The quotes ensure commas inside the fields are preserved.
            string[] csvLines =
            {
                "Name,Address",
                "\"Doe, John\",\"123 Main St, Apt 4\"",
                "\"Smith, Jane\",\"456 Oak Ave, Suite 12\""
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build a Word template programmatically with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("CSV LINQ Reporting Example");
            builder.Writeln("================================");
            // Begin a foreach loop over the CSV rows (exposed as a collection named 'persons').
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name   : <<[person.Name]>>");
            builder.Writeln("Address: <<[person.Address]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and prepare the CSV data source.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Configure CSV loading options: first line has headers, use double quotes for quoting.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                HasHeaders = true,
                QuoteChar = '"'
                // Delimiter defaults to ','; CommentChar defaults to '\0'.
            };

            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template is "persons".
            engine.BuildReport(loadedTemplate, csvDataSource, "persons");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
