using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with boolean values as true/false strings.
        string csvPath = "persons.csv";
        string[] csvLines =
        {
            "Name,IsActive",
            "Alice,true",
            "Bob,false",
            "Charlie,true"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create a Word template containing LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Active: <<[person.IsActive]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Configure CSV load options to recognize headers and default delimiter.
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true)
        {
            Delimiter = ',',
            QuoteChar = '"',
            // Boolean values represented as "true"/"false" are parsed automatically.
            // No additional configuration is required.
        };

        // Create a CSV data source using the configured options.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
