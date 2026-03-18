using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportGenerator
{
    static void Main()
    {
        // Path to the Word template that contains the report placeholders.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Path to the large CSV file that will be streamed.
        string csvPath = @"C:\Data\LargeData.csv";

        // Path where the generated report will be saved.
        string outputPath = @"C:\Reports\GeneratedReport.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Set up CSV parsing options (adjust as needed for your CSV format).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';      // Column separator.
        loadOptions.CommentChar = '#';    // Lines starting with this character are ignored.
        loadOptions.HasHeaders = true;   // First row contains column names.
        loadOptions.QuoteChar = '"';     // Character used to quote fields.

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
    }
}
