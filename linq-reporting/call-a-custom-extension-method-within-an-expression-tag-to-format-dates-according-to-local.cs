using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateTimeExtensions
{
    // Extension method that formats a DateTime according to the specified locale (culture name).
    public static string ToLocaleString(this DateTime date, string locale)
    {
        var culture = new CultureInfo(locale);
        // Use short date pattern for the culture.
        return date.ToString(culture.DateTimeFormat.ShortDatePattern, culture);
    }
}

// Simple data model with a DateTime property.
public class Order
{
    public DateTime OrderDate { get; set; } = DateTime.Now;
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting expressions that call the custom extension method.
        builder.Writeln("Order date (en-US): <<[order.OrderDate.ToLocaleString(\"en-US\")]>>");
        builder.Writeln("Order date (fr-FR): <<[order.OrderDate.ToLocaleString(\"fr-FR\")]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Create sample data.
        var order = new Order
        {
            // Use a fixed date for reproducibility.
            OrderDate = new DateTime(2023, 12, 25)
        };

        // Configure the reporting engine.
        var engine = new ReportingEngine();

        // Allow the engine to resolve extension methods.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Register the class that contains the extension method.
        engine.KnownTypes.Add(typeof(DateTimeExtensions));

        // Build the report. The root object name must match the tag prefix used in the template ("order").
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
