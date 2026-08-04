using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workDir, "data.csv");
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // Create a small CSV file with headers.
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        });

        // Build a template document containing LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("People List:");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure CSV data source options (first line has headers).
        var loadOptions = new CsvDataLoadOptions(true);
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Disable reflection optimization for this small data set.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
