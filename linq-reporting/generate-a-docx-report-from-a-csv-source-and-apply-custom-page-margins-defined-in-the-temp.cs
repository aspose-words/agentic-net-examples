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
        // Register code page provider for CSV encoding support (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths for the sample CSV, template, and final report.
        string csvPath = "sample.csv";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // Create sample CSV data.
        // -----------------------------------------------------------------
        File.WriteAllText(csvPath,
            "Name,Age,City\r\n" +
            "Alice,30,New York\r\n" +
            "Bob,25,London\r\n" +
            "Charlie,35,Sydney");

        // -----------------------------------------------------------------
        // Build the template document with custom page margins and LINQ tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Set custom margins (2 cm ≈ 56.7 points).
        builder.PageSetup.TopMargin = 56.7;
        builder.PageSetup.BottomMargin = 56.7;
        builder.PageSetup.LeftMargin = 56.7;
        builder.PageSetup.RightMargin = 56.7;

        // Title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Writeln("Customer Report");
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln();

        // Begin foreach block for the CSV rows.
        builder.Writeln("<<foreach [row in data]>>");
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row – each cell contains a LINQ Reporting tag.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.City]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options – only HasHeaders is required; defaults handle separator and encoding.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
