using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TaxReportModel
{
    // Sample data – initialize to avoid nullable warnings.
    public decimal Price { get; set; } = 100m;
    public decimal TaxRate { get; set; } = 0.07m; // 7 %
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare the data source.
        var model = new TaxReportModel();

        // 2. Create a blank Word document and insert LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a simple report layout.
        builder.Writeln("Price: <<[model.Price]>>");
        builder.Writeln("Tax Rate: <<[model.TaxRate]>>");
        // The expression below calculates the tax amount (price * taxRate).
        builder.Writeln("Calculated Tax: <<[model.Price * model.TaxRate]>>");

        // 3. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // The root object name must match the name used in the tags ("model").
        engine.BuildReport(doc, model, "model");

        // 4. Save the generated document.
        doc.Save("TaxReport.docx");
    }
}
