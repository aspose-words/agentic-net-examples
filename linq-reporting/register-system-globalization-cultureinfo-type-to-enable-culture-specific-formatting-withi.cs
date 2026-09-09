using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that format a date and a number using a specific culture.
        builder.Writeln("Date (fr-FR): <<[model.Date.ToString(\"D\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");
        builder.Writeln("Amount (fr-FR): <<[model.Amount.ToString(\"C\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel
        {
            Date = new DateTime(2023, 12, 31),
            Amount = 12345.67
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register System.Globalization.CultureInfo to allow its static members in template expressions.
        engine.KnownTypes.Add(typeof(CultureInfo));

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public DateTime Date { get; set; } = DateTime.Now;
    public double Amount { get; set; } = 0.0;
}
