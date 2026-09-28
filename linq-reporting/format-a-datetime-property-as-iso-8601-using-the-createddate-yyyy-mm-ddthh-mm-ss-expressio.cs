using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // Create a simple template with a LINQ Reporting tag that formats a DateTime as ISO‑8601
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated at: {=model.CreatedDate:yyyy-MM-ddTHH:mm:ss}");
        templateDoc.Save(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new()
        {
            CreatedDate = DateTime.Now
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(resultPath);

        // Indicate completion
        Console.WriteLine($"Report generated: {resultPath}");
    }

    public class ReportModel
    {
        public DateTime CreatedDate { get; set; }
    }
}
