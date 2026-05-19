using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Register code page provider for CSV parsing (required on some platforms).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllText(csvPath, "Name,Age\nJohn Doe,30\nJane Smith,25");

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        var loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };

        // Create the CSV data source.
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Disable reflection optimization for this report.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        doc.Save("report.docx");
    }
}
