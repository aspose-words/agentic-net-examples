using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        // Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that use CultureInfo for formatting.
        builder.Writeln("Date (en‑US): <<[model.Date.ToString(\"d\", CultureInfo.GetCultureInfo(\"en-US\"))]>>");
        builder.Writeln("Date (fr‑FR): <<[model.Date.ToString(\"d\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");
        builder.Writeln("Number (en‑US): <<[model.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"en-US\"))]>>");
        builder.Writeln("Number (fr‑FR): <<[model.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");

        // Save the template.
        string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Date = new DateTime(2023, 12, 31),
            Amount = 12345.67
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(CultureInfo)); // Register CultureInfo type.

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = System.IO.Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public DateTime Date { get; set; } = DateTime.Now;
    public double Amount { get; set; } = 0.0;
}
