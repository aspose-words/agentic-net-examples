using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that calls a static utility method.
        // The static method is accessed via the type name (Utility.FormatDate).
        builder.Writeln("Order date: <<[Utility.FormatDate(OrderDate)]>>");

        // Save the template to a temporary file.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        doc.Save(templatePath);

        // Load the template back (demonstrates load step).
        Document template = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            OrderDate = new DateTime(2023, 12, 25)
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the utility class so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Utility));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        template.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Sample data model used by the template.
public class ReportModel
{
    // The date that will be formatted by the utility method.
    public DateTime OrderDate { get; set; } = DateTime.MinValue;
}

// Utility class containing a static method that will be called from the template.
public static class Utility
{
    // Formats a DateTime as a short date string.
    public static string FormatDate(DateTime date)
    {
        return date.ToString("d");
    }
}
