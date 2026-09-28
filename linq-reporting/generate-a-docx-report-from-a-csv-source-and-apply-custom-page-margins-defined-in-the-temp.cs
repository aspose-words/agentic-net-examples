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

        // Working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string csvPath = Path.Combine(workDir, "data.csv");
        string outputPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a CSV source file with sample data.
        // -----------------------------------------------------------------
        File.WriteAllText(csvPath,
@"Name,Age,City
Alice,30,New York
Bob,25,Los Angeles
Charlie,35,Chicago");

        // -----------------------------------------------------------------
        // 2. Build the Word template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (1.5 cm ≈ 42.52 points).
        builder.PageSetup.TopMargin = 42.52f;
        builder.PageSetup.BottomMargin = 42.52f;
        builder.PageSetup.LeftMargin = 42.52f;
        builder.PageSetup.RightMargin = 42.52f;

        // Title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("CSV Report");
        builder.Writeln(); // blank line

        // Begin foreach over CSV rows.
        builder.Writeln("<<foreach [row in Data]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row (repeated for each CSV record).
        builder.InsertCell();
        builder.Font.Bold = false;
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.City]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using CSV data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Configure CSV data source (has headers, default comma separator).
        var loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // The default separator is a comma, so no explicit Separator property is needed.
        };
        var csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "Data");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
