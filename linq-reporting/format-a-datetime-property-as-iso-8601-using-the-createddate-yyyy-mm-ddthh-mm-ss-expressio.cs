using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            CreatedDate = DateTime.UtcNow
        };

        // Create a template document with an expression tag that formats the date as ISO 8601.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated at: {=CreatedDate:yyyy-MM-ddTHH:mm:ss}");
        doc.Save(templatePath);

        // Load the template and build the report.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        templateDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class ReportModel
{
    public DateTime CreatedDate { get; set; }
}
