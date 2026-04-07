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

        // Prepare file paths in the current working directory.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string csvPath = Path.Combine(dataDir, "people.csv");
        string templatePath = Path.Combine(dataDir, "template.docx");
        string outputPath = Path.Combine(dataDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a CSV file with a custom quote character '^'.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Description",
            "^John Doe^,^He said \"Hello, World\"^",
            "^Jane Smith^,^She replied \"Goodbye!\"^"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document that uses LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Description: <<[person.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, but demonstrates the load‑save cycle).
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure CSV loading options with the custom quote character.
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            QuoteChar = '^',
            HasHeaders = true
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
