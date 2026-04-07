using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionDemo
{
    // Extension methods must be defined in a static class.
    public static class AmountExtensions
    {
        // Returns true if the amount is greater than the specified limit.
        public static bool IsHighValue(this decimal amount, decimal limit) => amount > limit;
    }

    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public decimal Amount { get; set; }
    }

    public class ReportModel
    {
        // Collection of orders that the template will iterate over.
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple template:
            //   - Iterate over the Orders collection.
            //   - Output customer name and amount.
            //   - Use the custom extension method IsHighValue to display a flag.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Amount: $<<[order.Amount]>>");
            // Call the static method directly (Aspose.Words LINQ Reporting supports static method calls).
            builder.Writeln("<<if [AmountExtensions.IsHighValue(order.Amount, 1000)]>>   *** High Value ***<</if>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new()
            {
                Orders = new()
                {
                    new Order { CustomerName = "Alice", Amount = 750m },
                    new Order { CustomerName = "Bob",   Amount = 1250m },
                    new Order { CustomerName = "Carol", Amount = 500m }
                }
            };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the static class that contains the extension method so the template can call it.
            engine.KnownTypes.Add(typeof(AmountExtensions));

            // Build the report using the model. The root object name must match the template reference.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report_Output.docx");
        }
    }
}
