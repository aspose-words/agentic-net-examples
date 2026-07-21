using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string csvPath = Path.Combine(Environment.CurrentDirectory, "data.csv");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");

        // Create a simple CSV file with headers and sample data.
        File.WriteAllText(csvPath, "Id,Name,Age\r\n1,John Doe,30\r\n2,Jane Smith,25\r\n3,Bob Johnson,40");

        // Build a Word template containing LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from CSV data:");
        builder.Writeln("<<foreach [rec in records]>>");
        builder.Writeln("Id: <<[rec.Id]>>, Name: <<[rec.Name]>>, Age: <<[rec.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template document for reporting.
        Document doc = new Document(templatePath);

        // Configure CSV loading options (the file has a header row).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true;

        // Create a CSV data source based on the file and options.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the CSV data source. The data source name must match the name used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "records");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
