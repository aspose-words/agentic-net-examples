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
        // Register code page provider for CSV encoding support
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Name,Age,City\nAlice,30,New York\nBob,25,Los Angeles\nCharlie,35,Chicago");

        // Create a DOCX template with custom page margins and LINQ Reporting tags
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (1 inch = 72 points)
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;

        // Title
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Customer Report");
        builder.Font.ClearFormatting();
        builder.Writeln();

        // Begin foreach block that will repeat the whole table for each row
        builder.Writeln("<<foreach [row in data]>>");

        // Start table inside foreach
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row (repeated for each CSV record)
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.City]>>");
        builder.EndRow();

        // End table and foreach
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Configure CSV data source
        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };
        var csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // Save the generated report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
