using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a sample CSV file with a comment line, headers, and a custom ';' separator.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        var csvContent = new StringBuilder();
        csvContent.AppendLine("# This is a comment line and will be ignored");
        csvContent.AppendLine("Name;Age;City");
        csvContent.AppendLine("Alice;30;New York");
        csvContent.AppendLine("Bob;25;London");
        File.WriteAllText(csvPath, csvContent.ToString());

        // Configure CSV load options: the file has headers, uses ';' as delimiter, and '#' as comment character.
        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            // In the current Aspose.Words version the property is named 'Delimiter' (char).
            Delimiter = ';',
            CommentChar = '#'
        };

        // Create a CSV data source based on the file and the configured options.
        var csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Build the template document programmatically and insert LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [person in data]>>");
        builder.Writeln("Name: <<[person.Name]>> | Age: <<[person.Age]>> | City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Generate the report using the CSV data source.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, csvDataSource, "data");

        // Save the resulting report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        templateDoc.Save(outputPath);
    }
}
