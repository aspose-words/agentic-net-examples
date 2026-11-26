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
        // Register code page provider for CSV encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "data.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // Create sample CSV data.
        File.WriteAllText(csvPath,
            "Id,Name,Quantity,Price\n" +
            "1,Apple,10,0.5\n" +
            "2,Banana,20,0.3\n" +
            "3,Cherry,15,1.2");

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Define a custom table style (simple placeholder style).
        Style tableStyle = doc.Styles.Add(StyleType.Table, "MyTableStyle");

        // Insert the foreach tag that will iterate over CSV rows.
        builder.Writeln("<<foreach [row in csv]>>");

        // Start the table inside the foreach block.
        Table table = builder.StartTable();

        // Header row (will be repeated for each row, but acceptable for this simple example).
        builder.InsertCell(); builder.Write("Id");
        builder.InsertCell(); builder.Write("Name");
        builder.InsertCell(); builder.Write("Quantity");
        builder.InsertCell(); builder.Write("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell(); builder.Writeln("<<[row.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[row.Price]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Apply the custom style to the table.
        table.Style = tableStyle;

        // Save the template.
        doc.Save(templatePath);

        // -------------------------------------------------
        // Generate the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Load CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, new CsvDataLoadOptions { HasHeaders = true });

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "csv");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
