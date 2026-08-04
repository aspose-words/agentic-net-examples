using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model classes
    public class Order
    {
        public string Product { get; set; } = "";
        public int Quantity { get; set; }

        public Order() { }

        public Order(string product, int quantity)
        {
            Product = product;
            Quantity = quantity;
        }
    }

    public class Customer
    {
        public string Name { get; set; } = "";
        public List<Order> Orders { get; set; } = new();

        public Customer() { }

        public Customer(string name, List<Order> orders)
        {
            Name = name;
            Orders = orders;
        }
    }

    // Wrapper model containing the collection used in the template
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer("Alice", new List<Order>
                    {
                        new Order("Apple", 5),
                        new Order("Banana", 3)
                    }),
                    new Customer("Bob", new List<Order>
                    {
                        new Order("Carrot", 7),
                        new Order("Dates", 2),
                        new Order("Eggplant", 4)
                    })
                }
            };

            // Create the template document programmatically
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Outer foreach over customers
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln("Orders:");
            // Inner foreach over orders of the current customer
            builder.Writeln("<<foreach [order in customer.Orders]>>");
            builder.Writeln("- <<[order.Product]>> (Qty: <<[order.Quantity]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Build the report using the LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model);

            // Save the generated report
            doc.Save("Report.docx");
        }
    }
}
