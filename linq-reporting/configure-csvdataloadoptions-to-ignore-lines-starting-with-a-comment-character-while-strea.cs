using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CsvLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths in the current working directory.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "Template.docx");
            string csvPath = Path.Combine(workDir, "Data.csv");
            string outputPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a simple LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // The template iterates over a collection named "persons".
            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Create a CSV file that contains comment lines.
            // -----------------------------------------------------------------
            string[] csvLines =
            {
                "# This line is a comment and should be ignored",
                "Name,Age",
                "John,30",
                "# Another comment line",
                "Jane,25",
                "Bob,40"
            };
            File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 3. Configure CsvDataLoadOptions to treat lines starting with '#'
            //    as comments while streaming the CSV file.
            // -----------------------------------------------------------------
            var loadOptions = new CsvDataLoadOptions(hasHeaders: true)
            {
                Delimiter = ',',
                CommentChar = '#'
            };

            // Open the CSV file as a stream and create a CsvDataSource with the options.
            using var csvStream = File.OpenRead(csvPath);
            var csvDataSource = new CsvDataSource(csvStream, loadOptions);

            // -----------------------------------------------------------------
            // 4. Load the template document and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the data source name "persons" as referenced in the template.
            engine.BuildReport(reportDoc, csvDataSource, "persons");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
