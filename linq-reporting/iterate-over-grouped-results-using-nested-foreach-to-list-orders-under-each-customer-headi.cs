using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingNestedForeach
{
    // Data model classes
    public class Order
    {
        public int Id { get; set; }
        public string Product { get; set; } = "";
        public decimal Amount { get; set; }
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

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
                            new Order { Id = 1, Product = "Book", Amount = 12.5m },
                            new Order { Id = 2, Product = "Pen", Amount = 1.20m }
                        }
                    },
                    new Customer
                    {
                        Name = "Bob",
                        Orders = new List<Order>
                        {
                            new Order { Id = 3, Product = "Notebook", Amount = 5.00m }
                        }
                    }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Outer foreach: customers
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln("");

            // Inner foreach: orders for each customer
            builder.Writeln("<<foreach [order in customer.Orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>, Product: <<[order.Product]>>, Amount: <<[order.Amount]>>");
            builder.Writeln("<</foreach>>"); // end inner foreach

            builder.Writeln("<</foreach>>"); // end outer foreach

            // Save the template
            templateDoc.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
