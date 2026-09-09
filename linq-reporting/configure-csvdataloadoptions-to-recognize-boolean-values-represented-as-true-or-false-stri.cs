using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string csvPath = Path.Combine(workDir, "Data.csv");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // -------------------- Create template document --------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Person Report");
        // Use a foreach block to iterate over the CSV rows (treated as a collection named "persons")
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Active: <<[p.IsActive]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so it can be loaded later
        templateDoc.Save(templatePath);

        // -------------------- Create CSV data source --------------------
        // CSV content with a header row and boolean values as true/false strings
        string[] csvLines =
        {
            "Name,IsActive",
            "Alice,true",
            "Bob,false",
            "Charlie,true"
        };
        File.WriteAllLines(csvPath, csvLines);

        // Configure CSV load options: first line contains headers
        var loadOptions = new CsvDataLoadOptions(true);

        // Create the CSV data source using the configured options
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -------------------- Load template and build report --------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // The root object name "persons" matches the name used in the template tags
        engine.BuildReport(doc, csvDataSource, "persons");

        // -------------------- Save the generated report --------------------
        doc.Save(outputPath);
    }
}
