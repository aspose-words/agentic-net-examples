using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "Data.csv");
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25\nCharlie,35");

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over CSV rows (named "data").
        builder.Writeln("<<foreach [row in data]>>");

        // Start the table that will contain a header row and data rows.
        Table table = builder.StartTable();

        // Header row (static, appears once).
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row (repeated for each CSV record).
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // The default separator is a comma, so no explicit Separator property is needed.
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the final report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
