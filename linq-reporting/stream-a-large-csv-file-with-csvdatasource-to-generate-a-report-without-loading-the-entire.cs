using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsCsvStreamingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV parsing (required for some encodings).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(dataFolder);
            Directory.CreateDirectory(outputFolder);

            // Path to the large CSV file.
            string csvPath = Path.Combine(dataFolder, "large.csv");

            // Generate a large CSV file if it does not already exist.
            if (!File.Exists(csvPath))
            {
                using (var writer = new StreamWriter(csvPath))
                {
                    writer.WriteLine("Id,Name,Value"); // Header row.
                    for (int i = 0; i < 10000; i++)
                    {
                        writer.WriteLine($"{i},Name{i},{i * 10}");
                    }
                }
            }

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("CSV Report");
            // Iterate over the CSV rows (data source name will be "rows").
            builder.Writeln("<<foreach [row in rows]>>");
            builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Value: <<[row.Value]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport).
            string templatePath = Path.Combine(dataFolder, "template.docx");
            template.Save(templatePath);

            // Load the template document.
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Stream the CSV file using CsvDataSource (no full load into memory).
            // -----------------------------------------------------------------
            using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // Configure CSV loading options: first line contains headers.
                CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
                loadOptions.Delimiter = ','; // Default delimiter, set explicitly for clarity.

                // Create the CSV data source from the stream.
                CsvDataSource csvDataSource = new CsvDataSource(csvStream, loadOptions);

                // Build the report using the reporting engine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(reportDoc, csvDataSource, "rows");
            }

            // Save the generated report.
            string reportPath = Path.Combine(outputFolder, "Report.docx");
            reportDoc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine("Report generated successfully at:");
            Console.WriteLine(reportPath);
        }
    }
}
