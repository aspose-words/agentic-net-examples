using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class PriceReport
{
    // List of product prices.
    public List<decimal> Prices { get; set; } = new();

    // Returns the lowest price in the list.
    public decimal DiscountBenchmark => Prices.Min();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new PriceReport();
        model.Prices.AddRange(new[] { 199.99m, 149.50m, 179.75m, 129.99m, 159.00m });

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        // Insert a tag that references the pre‑computed lowest price.
        builder.Writeln("Discount Benchmark: <<[model.DiscountBenchmark]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
