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

        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a small CSV file with headers.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,Country",
            "Alice,30,USA",
            "Bob,25,Canada",
            "Charlie,35,UK"
        });

        // 2. Build a template document containing LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach loop over the CSV rows (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        // Output fields from each row.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Country: <<[person.Country]>>");
        builder.Writeln(); // Blank line between records.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Disable reflection optimization for this small CSV processing.
        ReportingEngine.UseReflectionOptimization = false;

        // 5. Configure CSV data source options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 6. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // 7. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(reportPath);

        // The example finishes without waiting for user input.
    }
}
