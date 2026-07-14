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

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllText(csvPath, "Name,Age\nJohn,30\nJane,25");

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure CSV data source options.
        var loadOptions = new CsvDataLoadOptions(hasHeaders: true);
        loadOptions.Delimiter = ',';
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Disable reflection optimization for this small data set.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
