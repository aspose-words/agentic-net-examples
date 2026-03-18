using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains LINQ Reporting Engine tags.
            string templatePath = @"C:\Docs\Template.docx";

            // Path to the large CSV file that will be used as a data source.
            string csvPath = @"C:\Data\LargeDataSet.csv";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Docs\Report.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Enable reflection optimization for faster property access on large CSV datasets.
            // This static property affects all ReportingEngine instances.
            ReportingEngine.UseReflectionOptimization = true;

            // Configure CSV parsing options (adjust as needed for your CSV format).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
            {
                // Assume the first line contains column headers.
                HasHeaders = true,
                // Use comma as the default delimiter; change if your CSV uses a different character.
                Delimiter = ',',
                // Optional: set comment and quote characters if required.
                CommentChar = '#',
                QuoteChar = '"'
            };

            // Create a CSV data source using the specified options.
            CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with data from the CSV source.
            // The second parameter is the data source object; the third parameter is the name
            // used to reference the data source inside the template (optional, can be null).
            engine.BuildReport(doc, dataSource, "Data");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
