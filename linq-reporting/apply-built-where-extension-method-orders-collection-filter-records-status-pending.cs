using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsExample
{
    // Simple POCO representing an order.
    public class Order
    {
        public int Id { get; set; }
        public string Status { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank Word document.
            Document doc = new Document();

            // 2. Prepare a sample collection of orders.
            List<Order> orders = new List<Order>
            {
                new Order { Id = 1, Status = "Pending" },
                new Order { Id = 2, Status = "Shipped" },
                new Order { Id = 3, Status = "Pending" },
                new Order { Id = 4, Status = "Cancelled" }
            };

            // 3. Use LINQ's Where extension method to filter orders with status "Pending".
            List<Order> pendingOrders = orders
                .Where(o => string.Equals(o.Status, "Pending", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // 4. Insert the filtered results into the Word document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Pending Orders:");
            foreach (Order order in pendingOrders)
            {
                builder.Writeln($"Order ID: {order.Id}, Status: {order.Status}");
            }

            // 5. Save the document to disk.
            doc.Save("PendingOrders.docx");
        }
    }
}
