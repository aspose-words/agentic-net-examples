using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Set the current culture to demonstrate locale‑specific number formatting.
        CultureInfo.CurrentCulture = new CultureInfo("de-DE");

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template with a LINQ Reporting tag that formats a number.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document templateDoc = new Document(templatePath);

        // Prepare the data model.
        Order order = new()
        {
            Total = 12345.67m
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(templateDoc, order, "order");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        templateDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath}");
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading.
        builder.Writeln("Invoice");

        // Insert a line that displays the total amount formatted according to the current culture.
        // The expression uses ToString with the "N" format specifier, which respects CultureInfo.CurrentCulture.
        builder.Writeln("Total Amount: <<[order.Total.ToString(\"N\")]>>");

        // Save the template.
        doc.Save(path);
    }
}

// Public data model class.
public class Order
{
    // Initialize to avoid nullable warnings.
    public decimal Total { get; set; } = 0m;
}
