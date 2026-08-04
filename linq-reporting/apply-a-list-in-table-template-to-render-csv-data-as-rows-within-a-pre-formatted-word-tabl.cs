using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\nCharlie,35");

        // Create a Word template with a pre‑formatted table and LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV load options – the first row contains column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true; // Enable header parsing so column names are recognised.

        // Use the CSV file as a data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report: the root data source name must match the tag reference ("data").
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the foreach block that will iterate over CSV rows.
        builder.Writeln("<<foreach [row in data]>>");

        // Create a table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – each cell contains a tag that references a CSV column.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}
