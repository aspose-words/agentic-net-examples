using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class CsvReportExample
{
    static void Main()
    {
        // Paths to the template document, CSV data file and the output report.
        string dataDir = @"C:\Data";
        string outputDir = @"C:\Output";
        string templatePath = Path.Combine(dataDir, "Template.docx");
        string csvPath = Path.Combine(dataDir, "ListOfPeople.csv");
        string outputPath = Path.Combine(outputDir, "Report.docx");

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure CSV loading options:
        // - The first line contains column headers.
        // - Use ';' as the column delimiter.
        // - Use '$' as the comment character.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ';';
        loadOptions.CommentChar = '$';

        // Create a CSV data source with the specified options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
