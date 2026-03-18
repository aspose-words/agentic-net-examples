using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Paths to the template, CSV file and output document.
        string myDir = @"C:\MyDir\";
        string artifactsDir = @"C:\ArtifactsDir\";

        // Load the Word template that contains the reporting tags.
        Document doc = new Document(myDir + "Template.docx");

        // Set up CSV parsing options.
        // - The CSV file has a header row (hasHeaders = true).
        // - Use a comma as the column delimiter.
        // - Use a single quote (') as the custom quote character to preserve quoted text exactly.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';
        loadOptions.HasHeaders = true;
        loadOptions.QuoteChar = '\''; // custom quote character

        // Create a CSV data source with the specified options.
        CsvDataSource dataSource = new CsvDataSource(myDir + "Data.csv", loadOptions);

        // Build the report by merging the template with the CSV data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save(artifactsDir + "Report.docx");
    }
}
