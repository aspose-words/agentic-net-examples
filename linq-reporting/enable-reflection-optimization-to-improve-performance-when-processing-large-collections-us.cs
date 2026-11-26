using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Sample data model
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Name = "Apple", Quantity = 3 },
            new Item { Name = "Banana", Quantity = 5 },
            new Item { Name = "Cherry", Quantity = 7 }
        };
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Enable reflection optimization for LINQ Reporting
            ReportingEngine.UseReflectionOptimization = true;

            // Create a new document and build the template with LINQ Reporting tags
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Prepare the data source
            Order order = new Order();

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report
            doc.Save("OrderReport.docx");
        }
    }
}
