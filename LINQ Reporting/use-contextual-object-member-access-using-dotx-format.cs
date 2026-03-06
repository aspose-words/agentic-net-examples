using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDotXExample
{
    // Simple data model to demonstrate DOTX member access.
    public class Order
    {
        public Customer Customer { get; set; }
        public List<OrderItem> Items { get; set; }
    }

    public class Customer
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }

    public class OrderItem
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write template text that uses DOTX syntax to access nested members.
            builder.Writeln("Customer Name: <<[order.Customer.Name]>>");
            builder.Writeln("Customer Email: <<[order.Customer.Email]>>");
            builder.Writeln("Order Items:");
            builder.Writeln("<<foreach [in order.Items]>><<[Product]>> - <<[Quantity]>>\n<</foreach>>");

            // Prepare sample data.
            Order sampleOrder = new Order
            {
                Customer = new Customer
                {
                    Name = "John Doe",
                    Email = "john.doe@example.com"
                },
                Items = new List<OrderItem>
                {
                    new OrderItem { Product = "Apple", Quantity = 3 },
                    new OrderItem { Product = "Banana", Quantity = 5 }
                }
            };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw if a member is not found.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Message to display for any missing member.
                MissingMemberMessage = "[Missing]"
            };

            // Build the report using the "order" object as a named data source.
            engine.BuildReport(doc, sampleOrder, "order");

            // Save the populated document.
            doc.Save("ReportWithDotX.docx");
        }
    }
}
