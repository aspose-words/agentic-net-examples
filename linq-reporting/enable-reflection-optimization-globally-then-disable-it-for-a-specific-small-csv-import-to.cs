using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Paths for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string templatePath = Path.Combine(workDir, "Template.docx");
        string csvPath = Path.Combine(workDir, "People.csv");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // Create a simple CSV file with a few rows.
        // -----------------------------------------------------------------
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "Alice,30\r\n" +
            "Bob,25\r\n" +
            "Charlie,35\r\n");

        // -----------------------------------------------------------------
        // Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Disable reflection optimization for this small CSV import.
        ReportingEngine.UseReflectionOptimization = false;

        // Prepare CSV data source.
        var csvOptions = new CsvDataLoadOptions(hasHeaders: true);
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
