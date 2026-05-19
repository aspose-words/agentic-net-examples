using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers and a few records.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Id,Name,Age",
            "1,John Doe,30",
            "2,Jane Smith,25",
            "3,Bob Johnson,40"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template that uses LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // List all persons using a foreach loop.
        builder.Writeln("All Persons:");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Display a specific record using ElementAt (zero‑based index).
        // Here we fetch the second person (index 1) from the collection.
        builder.Writeln("Selected Person (using ElementAt):");
        builder.Writeln("Name: <<[persons.ElementAt(1).Name]>>");
        builder.Writeln("Age: <<[persons.ElementAt(1).Age]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading options: first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ',';
        loadOptions.HasHeaders = true;

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // The example finishes without waiting for user input.
    }
}
