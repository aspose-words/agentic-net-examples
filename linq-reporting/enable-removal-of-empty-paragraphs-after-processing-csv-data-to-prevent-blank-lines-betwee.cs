using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string csvPath = Path.Combine(workDir, "Data.csv");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // 1. Create a CSV file. The second row is empty and would generate an empty paragraph.
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "John,30\r\n" +
            ",\r\n" +
            "Alice,25\r\n");

        // 2. Build a template document that contains LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>> Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Configure CSV loading options (header row present).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            HasHeaders = true,
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // 5. Create the CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report with the RemoveEmptyParagraphs option to delete blank lines.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(reportDoc, csvData, "persons");

        // 7. Save the final document.
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
