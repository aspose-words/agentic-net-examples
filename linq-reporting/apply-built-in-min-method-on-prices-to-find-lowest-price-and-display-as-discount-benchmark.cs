using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Prices = new List<decimal> { 199.99m, 149.50m, 179.75m, 129.00m }
        };

        // Create a blank document and insert a LINQ Reporting tag that references the minimum price.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Discount benchmark (lowest price): <<[model.MinPrice]>>");

        // Build the report using the model as the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("DiscountBenchmark.docx");
    }
}

// Public data model required by the LINQ Reporting engine.
public class ReportModel
{
    public List<decimal> Prices { get; set; } = new();

    // Expose the minimum price as a property so the template can access it.
    public decimal MinPrice => Prices.Min();
}
