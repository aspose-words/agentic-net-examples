using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define a working folder for all generated files.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workFolder);

        // -----------------------------------------------------------------
        // 1. Create a CSV file with a large number of rows (sample data).
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(workFolder, "SampleData.csv");
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Header row.
            writer.WriteLine("Id,Name,Value");

            // Generate 1000 sample rows.
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"{i},Item_{i},{(i * 0.5):F2}");
            }
        }

        // ---------------------------------------------------------------
        // 2. Build a Word template programmatically that uses LINQ tags.
        // ---------------------------------------------------------------
        string templatePath = Path.Combine(workFolder, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Add a simple title.
        builder.Writeln("Report generated from CSV data:");
        builder.Writeln();

        // LINQ Reporting tags:
        //   - The data source will be referenced by the name "data".
        //   - Iterate over each row and output its fields.
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and prepare the CSV data source.
        // ---------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        var loadOptions = new CsvDataLoadOptions(true);
        // Optional: customize delimiter, comment char, etc., if needed.
        // loadOptions.Delimiter = ',';
        // loadOptions.CommentChar = '#';

        // Create the CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // ---------------------------------------------------------------
        // 4. Enable reflection optimization for large data sets.
        // ---------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true;

        // ---------------------------------------------------------------
        // 5. Build the report using the ReportingEngine.
        // ---------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // ---------------------------------------------------------------
        // 6. Save the generated report.
        // ---------------------------------------------------------------
        string reportPath = Path.Combine(workFolder, "Report.docx");
        reportDoc.Save(reportPath);

        // Inform the user (no interactive prompts required).
        Console.WriteLine($"Report generated successfully: {reportPath}");
    }
}
