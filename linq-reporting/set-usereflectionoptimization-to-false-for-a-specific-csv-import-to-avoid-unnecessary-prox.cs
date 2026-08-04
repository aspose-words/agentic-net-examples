using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "people.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple CSV file with headers and sample data.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template that uses LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare CSV data source with header information.
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // CSV has headers.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 5. Disable reflection optimization for this specific import.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;

        // -----------------------------------------------------------------
        // 6. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
