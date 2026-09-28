using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with culture‑specific formatting tags
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Date (fr‑FR): <<[model.FormattedReportDate]>>");
        builder.Writeln("Total Amount (fr‑FR): <<[model.FormattedTotalAmount]>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new()
        {
            ReportDate = new DateTime(2023, 12, 31),
            TotalAmount = 12345.67
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string resultPath = Path.Combine(outputDir, "report.docx");
        doc.Save(resultPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {resultPath}");
    }
}

// Data model used by the template
public class ReportModel
{
    public DateTime ReportDate { get; set; } = DateTime.MinValue;
    public double TotalAmount { get; set; }

    // Culture‑specific formatted properties
    public string FormattedReportDate =>
        ReportDate.ToString("d", CultureInfo.GetCultureInfo("fr-FR"));

    public string FormattedTotalAmount =>
        TotalAmount.ToString("N2", CultureInfo.GetCultureInfo("fr-FR"));
}
