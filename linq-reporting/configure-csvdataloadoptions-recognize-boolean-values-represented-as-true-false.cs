using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvBooleanRecognitionExample
{
    static void Main()
    {
        // Path to the Word template that contains the reporting tags.
        string templatePath = @"C:\Docs\ReportTemplate.docx";

        // Path to the CSV file that contains boolean values as "true"/"false".
        string csvPath = @"C:\Docs\BooleanData.csv";

        // Path where the generated report will be saved.
        string outputPath = @"C:\Docs\GeneratedReport.docx";

        // Load the Word template.
        Document doc = new Document(templatePath);

        // Configure CSV loading options.
        // The constructor argument 'true' indicates that the first line contains column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Use the default comma delimiter (optional, shown for clarity).
        loadOptions.Delimiter = ',';
        // Ensure headers are recognized (already set by the constructor, kept for explicitness).
        loadOptions.HasHeaders = true;

        // Create a CSV data source with the specified load options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons"); // "persons" matches the root tag in the template.

        // Save the generated document.
        doc.Save(outputPath);
    }
}
