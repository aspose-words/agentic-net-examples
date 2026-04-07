using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample CSV file with headers and many rows.
        string csvPath = Path.Combine(workDir, "people.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Name,Age,Country");
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"Person {i},{20 + (i % 30)},{(i % 2 == 0 ? "USA" : "UK")}");
            }
        }

        // 2. Build a template document containing LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Write a heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach loop over the CSV rows (root name: persons).
        builder.Writeln("<<foreach [row in persons]>>");
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        builder.Writeln("Country: <<[row.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        var doc = new Document(templatePath);

        // 4. Enable reflection optimization for large data sets.
        ReportingEngine.UseReflectionOptimization = true;

        // 5. Configure CSV loading options (headers are present).
        var loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };

        // 6. Create a CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 7. Build the report using the data source and root name "persons".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // 8. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine("Report generated at: " + reportPath);
    }
}
