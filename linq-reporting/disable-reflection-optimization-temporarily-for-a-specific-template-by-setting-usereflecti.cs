using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Sample property used in the template.
    public string CustomerName { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Prepare the data source.
        Order order = new Order();

        // Store the original reflection‑optimization setting.
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;

        // Disable reflection optimization for this report generation.
        ReportingEngine.UseReflectionOptimization = false;

        try
        {
            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, order, "order");
        }
        finally
        {
            // Restore the original setting.
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }

        // Save the generated report.
        template.Save("Report.docx");
    }
}
