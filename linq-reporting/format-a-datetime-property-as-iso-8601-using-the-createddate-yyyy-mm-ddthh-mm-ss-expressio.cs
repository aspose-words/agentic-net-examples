using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize with UTC now to avoid nullable warnings.
    public DateTime CreatedDate { get; set; } = DateTime.UtcNow;
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Use a valid LINQ Reporting expression that formats the date via ToString().
        builder.Writeln("Report generated at: <<[model.CreatedDate.ToString(\"yyyy-MM-ddTHH:mm:ss\")]>>");

        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
