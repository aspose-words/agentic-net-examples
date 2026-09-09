using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();

        // Sample data bootstrap.
        public static ReportModel CreateSample()
        {
            var model = new ReportModel();

            var customer1 = new Customer { Name = "Alice Johnson" };
            customer1.Orders.Add(new Order { Product = "Laptop", Quantity = 1 });
            customer1.Orders.Add(new Order { Product = "Mouse", Quantity = 2 });

            var customer2 = new Customer { Name = "Bob Smith" };
            customer2.Orders.Add(new Order { Product = "Desk Chair", Quantity = 1 });

            model.Customers.Add(customer1);
            model.Customers.Add(customer2);

            return model;
        }
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public string Product { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Outer foreach over customers.
            builder.Writeln("<<foreach [customer in Customers]>>");
            builder.Writeln("Customer: <<[customer.Name]>>");
            builder.Writeln("Orders:");
            // Inner foreach over orders of the current customer.
            builder.Writeln("<<foreach [order in customer.Orders]>>");
            builder.Writeln("- <<[order.Product]>> (Qty: <<[order.Quantity]>>)"); 
            builder.Writeln("<</foreach>>"); // End inner foreach
            builder.Writeln("<</foreach>>"); // End outer foreach

            // 2. Save the template to a temporary file (required by the workflow).
            const string templatePath = "ReportTemplate.docx";
            template.Save(templatePath);

            // 3. Load the template (simulating a real‑world scenario where the template might be stored).
            Document doc = new Document(templatePath);

            // 4. Prepare the data source.
            ReportModel model = ReportModel.CreateSample();

            // 5. Build the report using Aspose.Words LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple example.
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            const string outputPath = "ReportResult.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
