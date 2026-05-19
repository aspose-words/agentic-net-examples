using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: format a date and a number using specific cultures.
        builder.Writeln("Date (French): <<[model.OrderDate.ToString(\"d\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");
        builder.Writeln("Number (German): <<[model.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"de-DE\"))]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel
        {
            OrderDate = new DateTime(2023, 12, 31),
            Amount = 1234567.89
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register System.Globalization.CultureInfo so it can be used inside template expressions.
        engine.KnownTypes.Add(typeof(CultureInfo));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public DateTime OrderDate { get; set; } = DateTime.Now;
    public double Amount { get; set; } = 0.0;
}
