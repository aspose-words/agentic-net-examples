using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedForeach
{
    // Data model classes
    public class Order
    {
        public int OrderId { get; set; } = 0;
        public string Product { get; set; } = "";
        public double Amount { get; set; } = 0.0;
    }

    public class Customer
    {
        public string Name { get; set; } = "";
        public List<Order> Orders { get; set; } = new();
    }

    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer
                    {
                        Name = "Alice Johnson",
                        Orders = new List<Order>
                        {
                            new Order { OrderId = 101, Product = "Laptop", Amount = 1299.99 },
                            new Order { OrderId = 102, Product = "Mouse", Amount = 25.50 }
                        }
                    },
                    new Customer
                    {
                        Name = "Bob Smith",
                        Orders = new List<Order>
                        {
                            new Order { OrderId = 201, Product = "Desk Chair", Amount = 199.99 },
                            new Order { OrderId = 202, Product = "Monitor", Amount = 349.00 },
                            new Order { OrderId = 203, Product = "Keyboard", Amount = 45.75 }
                        }
                    }
                }
            };

            // 2. Create a template document with LINQ Reporting tags
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Outer foreach for customers
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln("Orders:");

            // Inner foreach for orders of the current customer
            builder.Writeln("<<foreach [order in customer.Orders]>>");
            builder.Writeln("- Order ID: <<[order.OrderId]>>, Product: <<[order.Product]>>, Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>"); // end inner foreach

            builder.Writeln("<</foreach>>"); // end outer foreach

            // Save the template
            doc.Save(templatePath);

            // 3. Load the template and build the report
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 4. Save the generated report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Inform the user (no interactive input required)
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
