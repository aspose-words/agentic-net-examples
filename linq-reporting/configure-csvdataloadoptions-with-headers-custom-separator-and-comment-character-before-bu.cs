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

        // Prepare sample CSV data.
        string csvFile = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        string[] csvLines =
        {
            "# This is a comment line and will be ignored",
            "Id;Name;Score",
            "1;Alice;85",
            "2;Bob;92",
            "3;Charlie;78"
        };
        File.WriteAllLines(csvFile, csvLines, Encoding.UTF8);

        // Configure CSV load options: headers, custom separator ';', comment character '#'.
        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            Delimiter = ';',
            CommentChar = '#'
        };

        // Create CSV data source.
        var csvDataSource = new CsvDataSource(csvFile, csvOptions);

        // Build a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Score: <<[row.Score]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "data");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CsvReport.docx");
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
