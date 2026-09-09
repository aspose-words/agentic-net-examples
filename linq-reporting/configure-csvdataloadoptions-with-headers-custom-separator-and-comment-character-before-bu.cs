using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvReportExample
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllLines(csvPath, new[]
        {
            // Header line.
            "Name;Age;Country",
            // Data rows (using ';' as delimiter, '$' as comment character).
            "John Doe;30;USA",
            "$ This is a comment line and will be ignored",
            "Jane Smith;25;UK"
        });

        // Configure CSV loading options: headers present, custom delimiter ';', comment character '$'.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ';',
            CommentChar = '$',
            QuoteChar = '"'
        };

        // Create a CSV data source with the configured options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build a simple template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Country: <<[person.Country]>>");
        builder.Writeln("<</foreach>>");

        // Generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "persons");

        // Save the resulting document.
        template.Save("CsvReportOutput.docx");
    }
}
