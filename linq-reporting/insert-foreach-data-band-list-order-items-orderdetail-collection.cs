using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    // Simple data model representing an order with a collection of order details.
    public class Order
    {
        public int OrderId { get; set; }
        public List<OrderDetail> Details { get; set; } = new List<OrderDetail>();
    }

    public class OrderDetail
    {
        public string ItemName { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            Order sampleOrder = new Order { OrderId = 12345 };
            sampleOrder.Details.Add(new OrderDetail { ItemName = "Apple", Quantity = 3 });
            sampleOrder.Details.Add(new OrderDetail { ItemName = "Banana", Quantity = 5 });
            sampleOrder.Details.Add(new OrderDetail { ItemName = "Cherry", Quantity = 2 });

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a title for the order.
            builder.Writeln("Order Report");
            builder.Writeln($"Order ID: <<[order.OrderId]>>");
            builder.Writeln();

            // Insert a foreach data band that will iterate over Order.Details collection.
            // The syntax <<foreach [order.Details]>> starts the band,
            // and <</foreach>> ends it. Inside the band we reference fields of OrderDetail.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [order.Details]>>");
            builder.Writeln(" - Item: <<[ItemName]>>, Quantity: <<[Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "order" matches the prefix used in the template tags.
            engine.BuildReport(doc, sampleOrder, "order");

            // Save the resulting document.
            doc.Save("OrderReport.docx");
        }
    }
}
