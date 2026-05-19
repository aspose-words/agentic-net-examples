using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "Alice,30\r\n" +
            "Bob,25\r\n" +
            ",\r\n" +               // Empty row to produce empty paragraph after processing.
            "Charlie,35\r\n");

        // Create the template document.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static header.
        builder.Writeln("=== Report Header ===");
        builder.Writeln();

        // CSV section – will be populated via LINQ Reporting.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // This paragraph may become empty if the CSV row is empty.
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Static footer.
        builder.Writeln();
        builder.Writeln("=== Report Footer ===");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure CSV data source options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
        {
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // Create CSV data source.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report with removal of empty paragraphs.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, dataSource, "persons");

        // Save the final document.
        doc.Save("output.docx");
    }
}
