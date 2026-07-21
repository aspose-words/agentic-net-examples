using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvReportExample
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a CSV file with a custom quote character (single quote) and sample data.
        string csvPath = Path.Combine(workDir, "sample.csv");
        // Header line.
        string[] csvLines =
        {
            "Id,Description",
            "1,'Hello, World'",
            "2,'\"Quoted\" text'"
        };
        File.WriteAllLines(csvPath, csvLines);

        // 2. Create a template document programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a heading.
        builder.Writeln("CSV Report");
        builder.Writeln();

        // Begin a foreach block that iterates over the CSV rows.
        // The CSV data source will be referenced by the name "data".
        builder.Writeln("<<foreach [row in data]>>");
        // Output the Id and Description fields exactly as they appear in the CSV.
        builder.Writeln("Id: <<[row.Id]>>");
        builder.Writeln("Description: <<[row.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Configure CSV loading options with a custom quote character.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';      // Use comma as the column separator.
        loadOptions.QuoteChar = '\'';     // Use single quote as the quoting character.
        loadOptions.HasHeaders = true;    // First line contains column names.

        // 5. Create a CsvDataSource from the CSV file using the specified options.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // The root object name used in the template tags is "data".
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // 7. Save the generated report.
        string reportPath = Path.Combine(workDir, "CsvReport.docx");
        reportDoc.Save(reportPath);
    }
}
