using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        });

        // Create a template document with custom page margins.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Set custom margins (1 inch = 72 points).
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;

        // Insert LINQ Reporting tags to iterate over CSV rows.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("City: <<[p.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.HasHeaders = true;
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        bool success = engine.BuildReport(reportDoc, csvData, "persons");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to: {reportPath}");
    }
}
