using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;   // CsvDataSource and CsvDataLoadOptions are in this namespace

class CsvReportGenerator
{
    static void Main()
    {
        // Load the DOCX template that contains the custom table style.
        Document template = new Document("Template.docx");

        // Configure CSV parsing options.
        // The CSV file has a header row and uses a comma as the delimiter.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions(true);
        csvOptions.Delimiter = ',';      // column separator
        csvOptions.HasHeaders = true;    // first line contains column names

        // Create a CSV data source from the file.
        CsvDataSource csvSource = new CsvDataSource("Data.csv", csvOptions);

        // Populate the template with data from the CSV source.
        // The data source name ("Data") must match the name used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, csvSource, "Data");

        // Convert any table‑style formatting to direct formatting.
        // This guarantees that the custom style is applied to the generated tables.
        template.ExpandTableStylesToDirectFormatting();

        // Save the populated document.
        template.Save("Report.docx");
    }
}
