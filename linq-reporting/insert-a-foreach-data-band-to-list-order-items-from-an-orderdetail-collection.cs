using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<OrderDetail> OrderDetails { get; set; } = new();
    }

    public class OrderDetail
    {
        public string ProductName { get; set; } = "";
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create the template document programmatically
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln(); // empty line

            // Start a foreach data band that iterates over OrderDetails collection
            builder.Writeln("<<foreach [detail in OrderDetails]>>");
            builder.Writeln("Product: <<[detail.ProductName]>> , Quantity: <<[detail.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "OrderTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template for report generation
            var doc = new Document(templatePath);

            // 3. Prepare sample data
            var order = new Order
            {
                CustomerName = "John Doe",
                OrderDetails = new List<OrderDetail>
                {
                    new OrderDetail { ProductName = "Apple", Quantity = 5 },
                    new OrderDetail { ProductName = "Banana", Quantity = 3 },
                    new OrderDetail { ProductName = "Orange", Quantity = 7 }
                }
            };

            // 4. Build the report using ReportingEngine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // 5. Save the generated report
            const string outputPath = "OrderReport.docx";
            doc.Save(outputPath);
        }
    }
}
