using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // ---------- Create sample CSV data ----------
        string csvPath = Path.Combine(workDir, "people.csv");
        using (var writer = new StreamWriter(csvPath))
        {
            // Header.
            writer.WriteLine("Name,Age");
            // Sample rows (simulating a larger dataset).
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"Person {i},{20 + (i % 30)}");
            }
        }

        // ---------- Create a template document ----------
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tags: iterate over the CSV rows (exposed as 'persons').
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Load the template for reporting ----------
        var doc = new Document(templatePath);

        // ---------- Configure CSV data source ----------
        var loadOptions = new CsvDataLoadOptions(true) // CSV has headers.
        {
            Delimiter = ',',
            CommentChar = '#',
            QuoteChar = '"'
        };
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // ---------- Enable reflection optimization ----------
        ReportingEngine.UseReflectionOptimization = true;

        // ---------- Build the report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // ---------- Save the generated report ----------
        string reportPath = Path.Combine(workDir, "Report.docx");
        doc.Save(reportPath);
    }
}
