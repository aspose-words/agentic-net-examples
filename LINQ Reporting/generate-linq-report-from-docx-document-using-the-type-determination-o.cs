using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Sample enumeration that will be used in the template.
    public enum OrderStatus
    {
        Pending,
        Processing,
        Completed,
        Cancelled
    }

    // Data source class that contains an enum property and a collection.
    public class OrderReport
    {
        public OrderStatus Status { get; set; }
        public List<OrderItem> Items { get; set; }
    }

    public class OrderItem
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains LINQ Reporting Engine tags.
            // Example tag in the template: <<foreach [in ds.Items]>><<[Product]>><<[Quantity]>><<[Price]>><</foreach>>
            // Example enum usage: <<[ds.Status]:enumText>>
            Document doc = new Document("Template.docx");

            // Prepare the data source.
            var reportData = new OrderReport
            {
                Status = OrderStatus.Processing,
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Apple",  Quantity = 10, Price = 0.5m },
                    new OrderItem { Product = "Banana", Quantity = 5,  Price = 0.3m },
                    new OrderItem { Product = "Cherry", Quantity = 20, Price = 0.2m }
                }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the enum type so that the template can resolve it.
            engine.KnownTypes.Add(typeof(OrderStatus));

            // Build the report. The third parameter is the name used in the template to reference the data source.
            engine.BuildReport(doc, reportData, "ds");

            // Save the populated document.
            doc.Save("Report.docx");
        }
    }
}
