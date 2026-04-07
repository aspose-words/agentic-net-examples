using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with customer data
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(workDir, "customers.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Email,Address",
            "Alice Johnson,alice@example.com,123 Maple Street",
            "Bob Smith,bob@example.com,456 Oak Avenue",
            "Carol Davis,carol@example.com,789 Pine Road"
        });

        // -----------------------------------------------------------------
        // 2. Create a Word template that uses LINQ Reporting tags
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV data source (named "customers")
        builder.Writeln("<<foreach [c in customers]>>");
        builder.Writeln("--------------------------------------------------");
        builder.Writeln("Customer Report");
        builder.Writeln("Name   : <<[c.Name]>>");
        builder.Writeln("Email  : <<[c.Email]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln("--------------------------------------------------");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading options (header row present, comma delimiter)
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#',
            HasHeaders = true
        };

        // Create the CSV data source
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "customers");

        // -----------------------------------------------------------------
        // 4. Save the generated report
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "AllCustomersReport.docx");
        doc.Save(reportPath);

        // Inform the user (no interactive input required)
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
