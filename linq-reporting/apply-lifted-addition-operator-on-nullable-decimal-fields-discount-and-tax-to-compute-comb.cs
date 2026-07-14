using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data with nullable decimal fields.
        var order = new Order
        {
            Discount = 5.25m,
            Tax = 2.75m
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Insert a LINQ Reporting tag that outputs the combined value.
        builder.Writeln("Combined value: <<[order.Combined]>>");
        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for reporting.
        var loadedDoc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, order, "order");

        // Save the generated report.
        loadedDoc.Save("Report.docx");
    }
}

// Data model with nullable decimal fields and a computed property using lifted addition.
public class Order
{
    public decimal? Discount { get; set; }
    public decimal? Tax { get; set; }

    // The lifted addition operator returns null if either operand is null.
    public decimal? Combined => Discount + Tax;
}
