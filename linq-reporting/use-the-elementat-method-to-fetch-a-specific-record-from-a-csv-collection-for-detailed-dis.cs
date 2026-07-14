using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingElementAtExample
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory
        string csvPath = "people.csv";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // 1. Create a sample CSV file with headers and a few records
        File.WriteAllText(csvPath,
            "Id,Name,Age\r\n" +
            "1,John Doe,30\r\n" +
            "2,Jane Smith,25\r\n" +
            "3,Bob Johnson,40\r\n",
            Encoding.UTF8);

        // 2. Build a Word template programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Detailed information for the second record (index 1):");
        builder.Writeln("Id:   <<[persons.ElementAt(1).Id]>>");
        builder.Writeln("Name: <<[persons.ElementAt(1).Name]>>");
        builder.Writeln("Age:  <<[persons.ElementAt(1).Age]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // 3. Load the template document (could also reuse the instance, but following load rule)
        Document doc = new Document(templatePath);

        // 4. Configure CSV data source options (CSV has a header row)
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Create the CSV data source
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "persons"
        engine.BuildReport(doc, dataSource, "persons");

        // 6. Save the generated report
        doc.Save(outputPath);

        // Optional: indicate completion (no interactive input)
        Console.WriteLine("Report generated: " + Path.GetFullPath(outputPath));
    }
}
