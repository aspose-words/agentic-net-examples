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
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string csvPath = Path.Combine(outputDir, "data.csv");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the template document with custom margins and LINQ tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (in points).
        builder.PageSetup.TopMargin = 72;    // 1 inch
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 50;  // ~0.7 inch
        builder.PageSetup.RightMargin = 50;

        // Title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("People Report");
        builder.ParagraphFormat.ClearFormatting();

        // Begin foreach loop over CSV rows (root name: persons).
        builder.Writeln("<<foreach [p in persons]>>");

        // Table with header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row using tags.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.City]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';

        // Create CSV data source.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "persons");

        // Save the final report.
        reportDoc.Save(resultPath);
    }
}
