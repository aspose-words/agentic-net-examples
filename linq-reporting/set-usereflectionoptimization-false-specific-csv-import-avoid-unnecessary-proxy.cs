using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains the reporting engine tags.
        Document doc = new Document("Template.docx");

        // Disable the reflection‑based optimization for this run.
        // This prevents the engine from generating dynamic proxy classes,
        // which can be beneficial for small data sets.
        ReportingEngine.UseReflectionOptimization = false;

        // Configure CSV parsing options (adjust as needed for your CSV file).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // first line contains headers
        loadOptions.Delimiter = ',';      // column separator
        loadOptions.CommentChar = '#';    // comment character
        loadOptions.QuoteChar = '"';      // quote character

        // Create a CSV data source using the options above.
        CsvDataSource dataSource = new CsvDataSource("Data.csv", loadOptions);

        // Build the report by merging the template with the CSV data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
