using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template uses LINQ Reporting tags.
        // Date formatted using French culture.
        builder.Writeln(
            "Date (French): <<[model.Date.ToString(\"D\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");
        // Number formatted using German culture.
        builder.Writeln(
            "Amount (German): <<[model.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"de-DE\"))]>>");

        // Prepare the data model.
        ReportModel model = new()
        {
            Date = new DateTime(2023, 12, 31),
            Amount = 12345.67
        };

        // Configure the reporting engine.
        ReportingEngine engine = new();
        // Register System.Globalization.CultureInfo so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(CultureInfo));

        // Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("Report.docx");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public DateTime Date { get; set; }
    public double Amount { get; set; }
}
