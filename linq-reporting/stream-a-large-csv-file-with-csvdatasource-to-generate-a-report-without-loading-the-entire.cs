using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template, CSV data, and output report.
        const string templatePath = "template.docx";
        const string csvPath = "data.csv";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample large CSV file with headers and many rows.
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Id,Name,Age");
            for (int i = 1; i <= 1000; i++)
            {
                writer.WriteLine($"{i},Person {i},{20 + (i % 30)}");
            }
        }

        // -----------------------------------------------------------------
        // 2. Build the Word template containing LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("CSV Data Report");
        builder.Writeln("<<foreach [row in CsvData]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back for report generation.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Create a CsvDataSource that streams the CSV file.
        // -----------------------------------------------------------------
        using (var csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            var csvOptions = new CsvDataLoadOptions
            {
                HasHeaders = true // Use default separator (comma), quote and comment characters.
            };

            var csvDataSource = new CsvDataSource(csvStream, csvOptions);

            // -----------------------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, csvDataSource, "CsvData");
        }

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
