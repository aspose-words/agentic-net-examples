using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of price values.
    public List<decimal> Prices { get; set; } = new();

    // Returns the lowest price – calculated once and used in the template.
    public decimal DiscountBenchmark => Prices.Min();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        model.Prices.AddRange(new[] { 199.99m, 149.50m, 179.75m, 129.99m });

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Insert a tag that references the pre‑computed lowest price.
        builder.Writeln("Discount Benchmark: <<[model.DiscountBenchmark]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
