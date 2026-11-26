using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample prices.
    public List<decimal> Prices { get; set; } = new()
    {
        19.99m,
        5.49m,
        12.30m,
        8.75m
    };

    // Computed discount benchmark (lowest price).
    public decimal DiscountBenchmark => Prices.Min();
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Price Report");
        builder.Writeln("==============");
        builder.Writeln("Discount Benchmark (lowest price): $<<[model.DiscountBenchmark]>>");
        builder.Writeln(); // Add an empty line.

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new(templatePath);

        // Prepare the data model.
        ReportModel model = new();

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
