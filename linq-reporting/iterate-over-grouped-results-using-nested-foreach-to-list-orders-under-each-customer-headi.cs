using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedForeach
{
    // Data models
    public class Order
    {
        public int OrderId { get; set; }
        public string Product { get; set; } = "";
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
            // Prepare sample data
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer
                    {
                        Name = "Alice",
                        Orders = new List<Order>
                        {
                            new Order { OrderId = 1001, Product = "Laptop" },
                            new Order { OrderId = 1002, Product = "Mouse" }
                        }
                    },
                    new Customer
                    {
                        Name = "Bob",
                        Orders = new List<Order>
                        {
                            new Order { OrderId = 2001, Product = "Keyboard" },
                            new Order { OrderId = 2002, Product = "Monitor" }
                        }
                    }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Orders Report");
            builder.Writeln();

            // Outer foreach for customers
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Customer: <<[c.Name]>>");
            builder.Writeln("Orders:");
            // Inner foreach for orders of each customer
            builder.Writeln("<<foreach [o in c.Orders]>>");
            builder.Writeln("- Order ID: <<[o.OrderId]>>, Product: <<[o.Product]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
