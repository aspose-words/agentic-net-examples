using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // Create CSV file with quoted fields that contain commas.
        string csvPath = Path.Combine(workDir, "data.csv");
        string[] csvLines =
        {
            "\"Name\",\"Address\",\"Notes\"",
            "\"John Doe\",\"123 Main St, Apt 4\",\"He said, \"\"Hello\"\"\"",
            "\"Jane Smith\",\"456 Oak Ave, Suite 5\",\"New line\r\nin notes\""
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Create Word template programmatically.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("CSV Data Report");
        builder.Writeln();

        // Begin foreach over CSV rows (root name will be "data").
        builder.Writeln("<<foreach [row in data]>>");

        // Create a table with header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Address");
        builder.InsertCell();
        builder.Writeln("Notes");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Address]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Notes]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure CSV data source options.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            QuoteChar = '"'
            // The default separator is a comma, which matches the CSV format.
        };

        // Load CSV data source.
        CsvDataSource dataSource = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "data");

        // Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        doc.Save(reportPath);
    }
}
