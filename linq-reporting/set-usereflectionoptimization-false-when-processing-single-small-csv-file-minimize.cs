using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    class CsvReportGenerator
    {
        static void Main()
        {
            // Disable reflection optimization for small CSV processing.
            // This avoids the overhead of dynamic class generation.
            ReportingEngine.UseReflectionOptimization = false;

            // Load the template document that contains reporting engine tags.
            // Replace "Template.docx" with the path to your template.
            Document doc = new Document("Template.docx");

            // Configure CSV loading options (e.g., the file has a header row).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ',';      // Use comma as delimiter.
            loadOptions.CommentChar = '#';    // Lines starting with '#' are comments.
            loadOptions.QuoteChar = '"';      // Standard CSV quoting.

            // Create a CSV data source from the file.
            // Replace "Data.csv" with the path to your CSV file.
            CsvDataSource dataSource = new CsvDataSource("Data.csv", loadOptions);

            // Build the report using the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data"); // "Data" is the root object name in the template.

            // Save the generated report.
            // Replace "Result.docx" with the desired output path.
            doc.Save("Result.docx");
        }
    }
}
