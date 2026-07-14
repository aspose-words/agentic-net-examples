using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;          // ReportingEngine, CsvDataSource, CsvDataLoadOptions

public class Program
{
    public static void Main()
    {
        // Set up working paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string csvPath = Path.Combine(workDir, "people.csv");
        string outputPath = Path.Combine(workDir, "result.docx");

        // 1. Create a simple CSV file.
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,",          // Age is empty – will generate an empty paragraph after processing.
            "Charlie,25"
        });

        // 2. Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Section 1 – static content.
        builder.Writeln("=== Header Section ===");
        builder.Writeln();

        // Section 2 – CSV driven content.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Section 3 – static content.
        builder.Writeln("=== Footer Section ===");

        // Save the template (required before loading for reporting).
        template.Save(templatePath);

        // 3. Load the template back for reporting.
        Document doc = new Document(templatePath);

        // 4. Prepare the CSV data source.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 5. First pass: process CSV section with removal of empty paragraphs.
        ReportingEngine csvEngine = new ReportingEngine();
        csvEngine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        csvEngine.BuildReport(doc, csvData, "persons");

        // 6. Second pass: run a dummy build without the RemoveEmptyParagraphs flag
        // to keep other sections untouched (they have no tags, so this is a no‑op).
        ReportingEngine finalEngine = new ReportingEngine();
        finalEngine.BuildReport(doc, new object(), "");

        // 7. Save the final document.
        doc.Save(outputPath);
    }
}
