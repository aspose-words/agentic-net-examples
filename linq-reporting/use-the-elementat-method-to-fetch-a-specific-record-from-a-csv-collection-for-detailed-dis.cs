using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV handling (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a sample CSV file that will serve as the data source.
        // -----------------------------------------------------------------
        const string csvPath = "sample.csv";
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Id,Name,Age");
            writer.WriteLine("1,John,30");
            writer.WriteLine("2,Jane,25");
            writer.WriteLine("3,Bob,40");
        }

        // -----------------------------------------------------------------
        // Build the Word template programmatically and embed LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // List all records.
        builder.Writeln("All Records:");
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Display a specific record using ElementAt (second row, zero‑based index 1).
        builder.Writeln("Selected Record (second row):");
        builder.Writeln("Name: <<[data.ElementAt(1).Name]>>");
        builder.Writeln("Age: <<[data.ElementAt(1).Age]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // Configure CSV loading options (first row contains headers).
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        // Load CSV data source.
        var csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvData, "data");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
