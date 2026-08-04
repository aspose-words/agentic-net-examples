using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string csvPath = Path.Combine(outputDir, "Data.csv");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Generate a large CSV file line by line (streaming, no full load).
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath))
        {
            // Write header.
            writer.WriteLine("Name,Age,Country");

            // Write many rows.
            for (int i = 1; i <= 5000; i++)
            {
                writer.WriteLine($"Person {i},{20 + (i % 30)},{(i % 2 == 0 ? "USA" : "UK")}");
            }
        }

        // ---------------------------------------------------------------
        // 2. Create a simple Word template with LINQ Reporting tags.
        // ---------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>, Country: <<[person.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template for report generation.
        // ---------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // ---------------------------------------------------------------
        // 4. Prepare CSV data source using a stream (no full file load).
        // ---------------------------------------------------------------
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        loadOptions.Delimiter = ',';

        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            var csvDataSource = new CsvDataSource(csvStream, loadOptions);

            // -----------------------------------------------------------
            // 5. Build the report using ReportingEngine.
            // -----------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, csvDataSource, "persons");
        }

        // ---------------------------------------------------------------
        // 6. Save the generated report.
        // ---------------------------------------------------------------
        reportDoc.Save(resultPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
