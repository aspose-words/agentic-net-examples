using System;
using System.Collections.Generic;
using System.Linq; // Needed for LINQ methods
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // List of product prices.
    public List<decimal> Prices { get; set; } = new();

    // Exposes the minimum price for the template.
    public decimal MinPrice => Prices?.Min() ?? 0m;
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that references the MinPrice property.
        builder.Writeln("Discount benchmark price: <<[model.MinPrice]>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Prices.AddRange(new[] { 19.99m, 24.50m, 15.75m, 29.99m });

        // Build the report using the model as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
