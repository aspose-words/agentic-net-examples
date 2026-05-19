using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string csvPath = Path.Combine(workDir, "people.csv");
        string outputPath = Path.Combine(workDir, "ReportOutput.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple LINQ Reporting template.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from CSV data:");
        // Begin a foreach loop over the CSV data source named "persons".
        builder.Writeln("<<foreach [p in persons]>>");
        // Output each row's fields.
        builder.Writeln("- <<[p.Name]>>: <<[p.Age]>> years old");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a small CSV file with headers.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 3. Disable reflection optimization for this small CSV import.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;

        // Load CSV data with appropriate options (has headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Load the previously saved template.
        Document reportDoc = new Document(templatePath);

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "persons");

        // Save the generated report.
        reportDoc.Save(outputPath);

        // (Optional) Re‑enable reflection optimization for subsequent operations.
        ReportingEngine.UseReflectionOptimization = true;
    }
}
