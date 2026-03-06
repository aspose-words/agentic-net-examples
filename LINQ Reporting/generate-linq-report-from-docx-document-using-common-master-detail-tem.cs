using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains the master‑detail tags.
        // The template should have a master tag <<[Orders]>> and a nested detail tag <<[Items]>>.
        Document doc = new Document("Template.docx");

        // Prepare the master‑detail data using LINQ.
        // The master collection is a list of orders; each order contains a collection of items.
        List<object> masterData = GetSampleOrders()
            .Select(o => new
            {
                o.OrderId,
                o.CustomerName,
                // The detail collection is projected as a property named "Items".
                Items = o.Items.Select(i => new
                {
                    i.ProductName,
                    i.Quantity
                }).ToList()
            })
            .Cast<object>()
            .ToList();

        // Build the report. The data source name ("Orders") must match the master tag in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, masterData, "Orders");

        // Save the populated document.
        doc.Save("Report.docx");
    }

    // Generates sample master‑detail data.
    static List<Order> GetSampleOrders()
    {
        return new List<Order>
        {
            new Order
            {
                OrderId = 1001,
                CustomerName = "John Doe",
                Items = new List<OrderItem>
                {
                    new OrderItem { ProductName = "Pen", Quantity = 10 },
                    new OrderItem { ProductName = "Notebook", Quantity = 5 }
                }
            },
            new Order
            {
                OrderId = 1002,
                CustomerName = "Jane Smith",
                Items = new List<OrderItem>
                {
                    new OrderItem { ProductName = "Pencil", Quantity = 20 },
                    new OrderItem { ProductName = "Eraser", Quantity = 2 }
                }
            }
        };
    }

    // Master entity.
    class Order
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; }
        public List<OrderItem> Items { get; set; }
    }

    // Detail entity.
    class OrderItem
    {
        public string ProductName { get; set; }
        public int Quantity { get; set; }
    }
}
