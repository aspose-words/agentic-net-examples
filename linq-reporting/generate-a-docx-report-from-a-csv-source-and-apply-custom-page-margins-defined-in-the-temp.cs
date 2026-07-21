using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvReportExample
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        string csvPath = Path.Combine(outputDir, "Data.csv");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        // Columns: Name, Age, City
        string[] csvLines =
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document with custom margins and LINQ tags.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (1.5 cm left/right, 2 cm top/bottom).
        // 1 cm = 28.35 points.
        double cm = 28.35;
        builder.PageSetup.LeftMargin = 1.5 * cm;
        builder.PageSetup.RightMargin = 1.5 * cm;
        builder.PageSetup.TopMargin = 2.0 * cm;
        builder.PageSetup.BottomMargin = 2.0 * cm;

        // Title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("CSV Data Report");
        builder.Font.ClearFormatting();
        builder.Writeln();

        // LINQ Reporting tags: iterate over rows of the CSV data source named "data".
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        builder.Writeln("City: <<[row.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line has headers).
        var loadOptions = new CsvDataLoadOptions(true);

        // Create the CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using ReportingEngine.
        var engine = new ReportingEngine { Options = ReportBuildOptions.None };

        // The data source name used in the template tags is "data".
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(resultPath);
    }
}
