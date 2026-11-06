using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with a LINQ Reporting tag that references the minimum price.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Discount benchmark: <<[MinPrice]>>");
        builder.Document.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new ReportModel
        {
            Prices = new List<decimal> { 19.99m, 24.50m, 15.75m, 30.00m }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Collection of prices.
    public List<decimal> Prices { get; set; } = new();

    // Minimum price exposed for the template.
    public decimal MinPrice => Prices?.Min() ?? 0m;
}
