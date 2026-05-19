using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // ReportingEngine, CsvDataSource, CsvDataLoadOptions

public class Program
{
    public static void Main()
    {
        // Enable support for code pages (required for some CSV encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folders.
        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "Data");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample CSV file with customer data.
        string csvPath = Path.Combine(dataDir, "customers.csv");
        File.WriteAllText(csvPath,
            "Name,Email,Address\r\n" +
            "Alice Johnson,alice@example.com,123 Maple St.\r\n" +
            "Bob Smith,bob@example.com,456 Oak Ave.\r\n" +
            "Carol Davis,carol@example.com,789 Pine Rd.\r\n",
            Encoding.UTF8);

        // 2. Build a template document programmatically.
        string templatePath = Path.Combine(dataDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [c in customers]>>");
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Email: <<[c.Email]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        // Insert a page break after each customer's section.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template.
        Document doc = new Document(templatePath);

        // 4. Configure CSV data source options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // first line has headers
        {
            Delimiter = ',',
            CommentChar = '#',
            QuoteChar = '"'
        };

        // 5. Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "customers");

        // 7. Save the generated report containing a personalized page for each customer.
        string resultPath = Path.Combine(outputDir, "AllCustomersReport.docx");
        doc.Save(resultPath);
    }
}
