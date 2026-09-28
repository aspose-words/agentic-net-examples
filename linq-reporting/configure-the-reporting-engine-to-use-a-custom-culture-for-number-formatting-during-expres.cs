using System;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample numeric value to demonstrate culture‑specific formatting.
    public decimal Price { get; set; } = 1234.56m;
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Tag that will output the numeric value using the current culture.
        builder.Writeln("Price: <<[model.Price]>>");

        // Prepare the data source.
        ReportModel model = new();

        // Set the thread culture to French (France) so that number formatting uses a comma.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
        Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-FR");

        // Build the report using the document template, the data source, and the root name "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
