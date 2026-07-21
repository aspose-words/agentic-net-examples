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

        // Prepare sample CSV data with a header row and boolean values as "true"/"false".
        string csvPath = "people.csv";
        string[] csvLines =
        {
            "Name,IsActive",
            "Alice,true",
            "Bob,false",
            "Charlie,true"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Configure CSV loading options:
        // - The first line contains column names (HasHeaders = true).
        // - Use the default comma delimiter.
        // - No comment or quote characters needed for this simple example.
        var loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ','; // optional, default is ','.
        loadOptions.HasHeaders = true;

        // Create a CSV data source using the configured options.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build a Word template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a simple report that iterates over the CSV rows.
        builder.Writeln("People Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Active: <<[person.IsActive]>>");
        builder.Writeln("<</foreach>>");

        // Create the reporting engine and generate the report.
        var engine = new ReportingEngine();
        // The data source name ("persons") must match the name used in the template tags.
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        string outputPath = "PeopleReport.docx";
        doc.Save(outputPath);
    }
}
