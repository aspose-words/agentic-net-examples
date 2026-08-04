using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample CSV file with headers and three records.
        string csvPath = Path.Combine(workDir, "people.csv");
        File.WriteAllText(csvPath,
            "Id,Name,Age\r\n" +
            "1,John Doe,30\r\n" +
            "2,Jane Smith,25\r\n" +
            "3,Bob Johnson,40\r\n");

        // 2. Build a Word template programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("=== CSV LINQ Reporting Example ===");
        builder.Writeln();

        // Display the third record (index 2) using ElementAt.
        builder.Writeln("Detailed view of the third record (ElementAt):");
        builder.Writeln("Name: <<[persons.ElementAt(2).Name]>>");
        builder.Writeln("Age:  <<[persons.ElementAt(2).Age]>>");
        builder.Writeln();

        // Optional: list all records using a foreach loop.
        builder.Writeln("All records:");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("- Id: <<[p.Id]>>, Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Prepare CSV data source with header support.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Build the report using the data source named "persons".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive input).
        Console.WriteLine("Report generated at: " + reportPath);
    }
}
