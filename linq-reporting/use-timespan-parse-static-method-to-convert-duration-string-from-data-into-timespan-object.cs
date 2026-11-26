using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Original duration string from data source.
    public string DurationString { get; set; } = "02:30:45";

    // Parsed TimeSpan using TimeSpan.Parse.
    public TimeSpan Duration => TimeSpan.Parse(DurationString);
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple line that will display the parsed TimeSpan.
        builder.Writeln("Duration: <<[order.Duration]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and prepare the data source.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        Order order = new Order(); // DurationString is already set.

        // -------------------------------------------------
        // Build the report using Aspose.Words LINQ Reporting Engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
