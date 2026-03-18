using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the CSV file.
        const string csvPath = "data.csv";

        // Create CSV content with quoted fields that contain commas.
        // Header row.
        // Data row: Name = "Doe, John", Address = "123 Main St, Apt 4"
        string[] csvLines =
        {
            "Name,Address",
            "\"Doe, John\",\"123 Main St, Apt 4\""
        };
        File.WriteAllLines(csvPath, csvLines);

        // Create a simple Word template in memory.
        // The template uses Reporting Engine syntax to reference the data source named "persons".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Name: <<[persons.Name]>>");
        builder.Writeln("Address: <<[persons.Address]>>");

        // Configure CSV parsing options.
        // HasHeaders = true tells the engine that the first row contains column names.
        // QuoteChar is set to double‑quote to correctly handle quoted fields.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';   // Default delimiter, set explicitly for clarity.
        loadOptions.QuoteChar = '\"'; // Default quote character.

        // Create the CSV data source using the file path and the load options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report by merging the template with the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
