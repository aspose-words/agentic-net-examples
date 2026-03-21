using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportGenerator
{
    static void Main()
    {
        // Use the executable's directory for all files so they always exist.
        string baseDir = AppContext.BaseDirectory;

        // Path to the Word template that contains the report placeholders.
        string templatePath = Path.Combine(baseDir, "ReportTemplate.docx");

        // Path to the large CSV file that will be streamed.
        string csvPath = Path.Combine(baseDir, "LargeData.csv");

        // Path where the generated report will be saved.
        string outputPath = Path.Combine(baseDir, "GeneratedReport.docx");

        // Ensure the template document exists.
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln($"Report generated on {DateTime.Now:yyyy-MM-dd}");
            builder.InsertField("MERGEFIELD Data.Name");
            builder.InsertParagraph();
            builder.InsertField("MERGEFIELD Data.Value");
            templateDoc.Save(templatePath);
        }

        // Ensure the CSV file exists (create a modestly sized sample).
        if (!File.Exists(csvPath))
        {
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                writer.WriteLine("Name,Value"); // header
                for (int i = 1; i <= 1000; i++)
                {
                    writer.WriteLine($"Item{i},{i}");
                }
            }
        }

        // Load the template document.
        Document doc = new Document(templatePath);

        // Set up CSV parsing options (adjust as needed for your CSV format).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            CommentChar = '#',
            HasHeaders = true,
            QuoteChar = '"'
        };

        // Open the CSV file as a read‑only stream; the file is never fully loaded into memory.
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            // Create a CSV data source that reads directly from the stream using the defined options.
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // Populate the template with data from the CSV source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data");

            // Save the resulting report.
            doc.Save(outputPath);
        }

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
