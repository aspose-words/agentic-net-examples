using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the template Word document that contains the reporting syntax.
        // The template should have a placeholder like <<[ds.Amount] + [ds.Tax]>> to display the sum.
        string templatePath = @"C:\Docs\ReportTemplate.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure CSV loading options.
        // HasHeaders = true tells the engine that the first line contains column names.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';          // Use comma as column separator.
        loadOptions.CommentChar = '#';        // Optional: lines starting with # are ignored.
        loadOptions.QuoteChar = '"';          // Optional: quote character for fields containing delimiters.

        // Path to the CSV file that provides the data source.
        // The CSV must contain columns named "Amount" and "Tax" (or any names you use in the template).
        string csvPath = @"C:\Data\FinancialData.csv";

        // Create the CSV data source using the file path and the load options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        // The third argument ("ds") is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds");

        // Save the generated report.
        string outputPath = @"C:\Docs\GeneratedReport.docx";
        doc.Save(outputPath);
    }
}
