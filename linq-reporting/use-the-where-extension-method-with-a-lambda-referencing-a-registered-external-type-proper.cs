using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public double Amount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// External type whose static property will be used in the LINQ expression.
public static class FilterHelper
{
    public static double MinAmount { get; set; } = 150;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tag using Where with a lambda that references the external static property.
        builder.Writeln("<<foreach [order in Orders.Where(o => o.Amount > FilterHelper.MinAmount)]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>> - Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before loading it for reporting).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data source.
        // -------------------------------------------------
        var model = new ReportModel();
        model.Orders.Add(new Order { CustomerName = "Alice", Amount = 120 });
        model.Orders.Add(new Order { CustomerName = "Bob", Amount = 200 });
        model.Orders.Add(new Order { CustomerName = "Charlie", Amount = 300 });

        // -------------------------------------------------
        // 4. Configure and execute the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();

        // Register the external type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(FilterHelper));

        // Build the report using the model as the root data source.
        engine.BuildReport(reportDoc, model);

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
