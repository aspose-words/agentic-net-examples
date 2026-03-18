using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains the reporting tags.
        // The template should have inline calculations, e.g. <<[ds.Amount]:sum>>.
        Document doc = new Document("Template.docx");

        // Configure CSV loading options – the first line contains column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';          // Use comma as the field separator.
        loadOptions.CommentChar = '#';        // Optional: lines starting with # are ignored.
        loadOptions.HasHeaders = true;        // First row is treated as column names.
        loadOptions.QuoteChar = '"';          // Optional: quote character for fields.

        // Create a CSV data source from the CSV file using the options above.
        CsvDataSource csvData = new CsvDataSource("Data.csv", loadOptions);

        // Build the report by merging the CSV data into the template.
        // The data source name "ds" is used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "ds");

        // Save the populated document.
        doc.Save("Report.docx");
    }
}
