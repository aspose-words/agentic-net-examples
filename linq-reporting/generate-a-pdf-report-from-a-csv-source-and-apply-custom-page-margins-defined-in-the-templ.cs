using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "data.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "Report.pdf");

        // Create a simple CSV file with headers.
        File.WriteAllText(csvPath, "Name,Age,Country\nAlice,30,USA\nBob,25,Canada\nCharlie,35,UK");

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Apply custom page margins (points; 1 inch = 72 points).
        builder.PageSetup.TopMargin = 72;      // 1 inch
        builder.PageSetup.BottomMargin = 72;   // 1 inch
        builder.PageSetup.LeftMargin = 50;     // ~0.7 inch
        builder.PageSetup.RightMargin = 50;    // ~0.7 inch

        // Add a title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("People Report");
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln(); // empty line

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Country: <<[person.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // Save the final report as PDF.
        reportDoc.Save(outputPath, SaveFormat.Pdf);
    }
}
