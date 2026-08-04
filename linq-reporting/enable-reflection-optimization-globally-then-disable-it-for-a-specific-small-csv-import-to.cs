using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a small CSV file.
        string csvPath = Path.Combine(workDir, "people.csv");
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines);

        // 2. Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // 5. Disable the optimization for this small CSV import to avoid overhead.
        ReportingEngine.UseReflectionOptimization = false;

        // 6. Prepare CSV data source with header handling.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 7. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // 8. Save the generated report.
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);

        // Inform that the process completed.
        Console.WriteLine($"Report generated at: {outputPath}");
    }
}
