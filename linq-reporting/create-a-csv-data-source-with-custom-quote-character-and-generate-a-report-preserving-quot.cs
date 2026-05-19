using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample CSV file.
        // The CSV uses double quotes to enclose a field that contains a comma and quotes.
        // To embed a quote inside a quoted field we double it.
        string csvPath = Path.Combine(outputDir, "SampleData.csv");
        string[] csvLines =
        {
            "Name,Comment",
            "Alice,\"\"\"Hello, World!\"\"\"",
            "Bob,No quotes"
        };
        File.WriteAllLines(csvPath, csvLines);

        // 2. Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a foreach loop to iterate over CSV rows.
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Comment: <<[row.Comment]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document reportDoc = new Document(templatePath);

        // 4. Configure CSV load options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
        {
            Delimiter = ',',
            QuoteChar = '"', // standard double‑quote character
            HasHeaders = true
        };

        // 5. Create CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // 7. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
