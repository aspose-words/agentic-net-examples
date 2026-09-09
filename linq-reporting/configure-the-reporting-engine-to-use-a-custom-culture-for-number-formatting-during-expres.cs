using System;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample numeric value that will be formatted according to the custom culture.
    public decimal Price { get; set; } = 1234.56m;
}

public class Program
{
    public static void Main()
    {
        // Set a custom culture (French) for the current thread.
        // This culture uses a comma as the decimal separator.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // Create the template document and insert a LINQ Reporting tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Price: <<[model.Price]>>");

        // Build the report using the custom culture.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new ReportModel(), "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
