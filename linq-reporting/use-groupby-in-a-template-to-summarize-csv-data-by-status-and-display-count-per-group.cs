using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingGroupByExample
{
    public static void Main()
    {
        // Register code page provider for CSV support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
@"Id,Name,Status
1,Apple,Available
2,Banana,OutOfStock
3,Cherry,Available
4,Date,Available
5,Elderberry,OutOfStock");

        // Build a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Summary by Status");
        builder.Writeln("");
        builder.Writeln("<<foreach [g in Data.GroupBy(d => d.Status)]>>");
        builder.Writeln("Status: <<[g.Key]>> - Count: <<[g.Count()]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Load CSV data source.
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "Data");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
