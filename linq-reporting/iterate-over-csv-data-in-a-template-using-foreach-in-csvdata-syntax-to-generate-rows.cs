using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        File.WriteAllText(csvPath, "Id,Name,Amount\n1,Apple,10.5\n2,Banana,20\n3,Cherry,15.75");

        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [row in csvData]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Amount: <<[row.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source options (only HasHeaders is supported in this version).
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Load CSV data source.
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "csvData");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
