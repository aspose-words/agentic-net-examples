using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of prices.
    public List<decimal> Prices { get; set; } = new();

    // Lowest price calculated using LINQ.
    public decimal DiscountBenchmark => Prices.Min();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Prices = new List<decimal> { 199.99m, 149.50m, 179.75m, 129.99m }
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Use the calculated property in the template.
        builder.Writeln("Discount benchmark price: <<[model.DiscountBenchmark]>>");
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
