using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Original duration string (e.g., "02:15:30")
    public string DurationString { get; set; } = string.Empty;

    // Parsed TimeSpan value
    public TimeSpan Duration { get; set; }

    public Order(string durationString)
    {
        DurationString = durationString;
        // Convert the string to a TimeSpan using the static Parse method
        Duration = TimeSpan.Parse(durationString);
    }
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple LINQ Reporting tag that will display the parsed TimeSpan
        builder.Writeln("Order duration: <<[order.Duration]>>");

        // Save the template to disk
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source
        // -----------------------------------------------------------------
        // Example duration string; you can change this to any valid TimeSpan format
        Order order = new Order("02:15:30"); // 2 hours, 15 minutes, 30 seconds

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting Engine
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name ("order") must match the tag prefix used in the template
        engine.BuildReport(reportDoc, order, "order");

        // -----------------------------------------------------------------
        // 5. Save the generated report
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
