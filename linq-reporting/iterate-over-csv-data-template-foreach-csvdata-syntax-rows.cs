using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the Word template that contains the <<foreach [in csvData]>> tag.
        string templatePath = @"C:\Docs\Template.docx";

        // Path to the CSV file that will be used as the data source.
        string csvPath = @"C:\Docs\People.csv";

        // Path where the generated document will be saved.
        string outputPath = @"C:\Docs\Result.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Set CSV parsing options. The first row of the CSV contains column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Example of customizing parsing (optional):
        // loadOptions.Delimiter = ';';
        // loadOptions.CommentChar = '#';
        // loadOptions.QuoteChar = '"';

        // Create a CsvDataSource instance from the CSV file using the options above.
        CsvDataSource csvSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report. The third argument ("csvData") is the name that must match the
        // identifier used in the template's foreach tag: <<foreach [in csvData]>>.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvSource, "csvData");

        // Save the populated document.
        doc.Save(outputPath);
    }
}
