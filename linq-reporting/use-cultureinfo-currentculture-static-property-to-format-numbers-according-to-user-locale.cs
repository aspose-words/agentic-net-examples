using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            Amount = 12345.67m
        };

        // Create a template document with a LINQ Reporting tag.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Order amount: <<[order.Amount]>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Set the current culture to demonstrate locale‑specific formatting.
        // For example, French (France) uses a comma as the decimal separator.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Public data model required by the template.
public class Order
{
    // Decimal values are formatted according to CultureInfo.CurrentCulture when rendered.
    public decimal Amount { get; set; } = 0m;
}
