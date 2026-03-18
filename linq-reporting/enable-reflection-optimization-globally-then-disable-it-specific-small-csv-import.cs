using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReflectionOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Enable reflection optimization globally – this is the default,
            // but we set it explicitly for clarity.
            ReportingEngine.UseReflectionOptimization = true;

            // Path to the template document that will receive the report data.
            string templatePath = @"C:\Docs\ReportTemplate.docx";

            // Path to the small CSV file that we want to import without the
            // reflection optimization overhead.
            string csvPath = @"C:\Data\SmallData.csv";

            // -----------------------------------------------------------------
            // Build a report for the small CSV while the optimization is disabled.
            // -----------------------------------------------------------------
            // Temporarily turn off the optimization.
            ReportingEngine.UseReflectionOptimization = false;

            // Configure CSV loading options (adjust as needed).
            CsvDataLoadOptions csvLoadOptions = new CsvDataLoadOptions(hasHeaders: true);
            csvLoadOptions.Delimiter = ',';          // Use comma as delimiter.
            csvLoadOptions.CommentChar = '#';        // Optional comment character.
            csvLoadOptions.QuoteChar = '"';          // Optional quote character.

            // Create a CSV data source based on the file and the load options.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, csvLoadOptions);

            // Load the template document.
            Document reportDoc = new Document(templatePath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name of the data source as referenced in the template.
            engine.BuildReport(reportDoc, csvDataSource, "CsvData");

            // Restore the global optimization setting for subsequent operations.
            ReportingEngine.UseReflectionOptimization = true;

            // Save the generated report.
            string outputPath = @"C:\Docs\ReportFromSmallCsv.docx";
            reportDoc.Save(outputPath);
        }
    }
}
