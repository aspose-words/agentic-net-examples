using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvReportExample
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with a custom single‑quote as the quote character.
        const string csvPath = "people.csv";
        File.WriteAllText(csvPath,
@"Id,Name,Comment
1,'Alice','""Hello, World""'
2,'Bob','""Goodbye, World""'");

        // Configure CSV loading options: headers present, comma delimiter, single‑quote as quote char.
        var loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '\''
        };

        // Create a CSV data source from the file using the custom options.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build a simple Word template programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Id: <<[person.Id]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Comment: <<[person.Comment]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template (demonstrating load step).
        var loadedTemplate = new Document(templatePath);

        // Build the report using the CSV data source.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, csvDataSource, "persons");

        // Save the generated report.
        const string reportPath = "report.docx";
        loadedTemplate.Save(reportPath);
    }
}
