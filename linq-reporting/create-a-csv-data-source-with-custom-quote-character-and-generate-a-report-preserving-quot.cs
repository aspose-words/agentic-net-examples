using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // File paths for CSV data, template, and output report.
        string dataFile = "data.csv";
        string templateFile = "template.docx";
        string outputFile = "report.docx";

        // Create CSV content with a custom quote character (backtick `).
        string[] csvLines = new[]
        {
            "Name,Description",
            "`John Doe`,`\"He said, \"\"Hello\"\" to everyone\"`",
            "`Jane Smith`,`\"She replied, \"\"Hi!\"\"\"`"
        };
        File.WriteAllLines(dataFile, csvLines);

        // Build the template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Description: <<[person.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to follow the load/save lifecycle.
        doc.Save(templateFile);
        Document template = new Document(templateFile);

        // Configure CSV load options with the custom quote character.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '`' // Custom quote character.
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(dataFile, loadOptions);

        // Generate the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, csvDataSource, "persons");

        // Save the final report.
        template.Save(outputFile);
    }
}
