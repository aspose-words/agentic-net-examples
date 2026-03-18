using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportExample
{
    static void Main()
    {
        // Load the Word template that contains the report placeholders.
        Document doc = new Document(@"C:\Templates\ReportTemplate.docx");

        // Configure CSV loading options.
        // Set HasHeaders to true if the first line contains column names.
        // Set CommentChar to the character that marks comment lines (e.g., '#').
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.CommentChar = '#';

        // Open the large CSV file as a stream to avoid loading the entire file into memory.
        using (FileStream csvStream = File.OpenRead(@"C:\Data\LargeDataFile.csv"))
        {
            // Create a CSV data source using the stream and the configured options.
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // Build the report by merging the data source with the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data");
        }

        // Save the generated report.
        doc.Save(@"C:\Output\ReportResult.docx");
    }
}
