using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Prepare the data model.
        Order order = new Order { CustomerName = "John Doe" };

        // Preserve the original reflection optimization setting.
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;

        try
        {
            // Disable reflection optimization for this report generation.
            ReportingEngine.UseReflectionOptimization = false;

            // Build the report using the template and the data model.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, order, "order");
        }
        finally
        {
            // Restore the original optimization setting.
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }

        // Save the generated report.
        template.Save("Report.docx");
    }

    // Simple data model class used by the template.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
    }
}
