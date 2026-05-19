using System;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Ensure the current thread uses the desired culture.
        Thread.CurrentThread.CurrentCulture = CultureInfo.CurrentCulture;

        // Step 1: Create the data model.
        var model = new ReportModel
        {
            Order = new Order
            {
                Id = 12345,
                CustomerName = "John Doe",
                Total = 1234.56m
            }
        };

        // Step 2: Build the template document programmatically.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("Order ID: <<[model.Order.Id]>>");
        builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
        // The LINQ Reporting engine will call ToString() on the decimal value,
        // which respects the current thread's culture set above.
        builder.Writeln("Total: <<[model.Order.Total]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Step 3: Load the template and generate the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 4: Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Public data model classes required by the LINQ Reporting engine.
public class ReportModel
{
    public Order Order { get; set; } = new Order();
}

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }
}
