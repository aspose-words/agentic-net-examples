using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model representing an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<OrderDetail> OrderDetails { get; set; } = new();
    }

    // Simple data model representing a line item in an order.
    public class OrderDetail
    {
        public string ProductName { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words for some encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                OrderDetails = new List<OrderDetail>
                {
                    new OrderDetail { ProductName = "Apple", Quantity = 3 },
                    new OrderDetail { ProductName = "Banana", Quantity = 5 },
                    new OrderDetail { ProductName = "Orange", Quantity = 2 }
                }
            };

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header with customer name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin a foreach data band that iterates over OrderDetails.
            builder.Writeln("Order Items:");
            builder.Writeln("<<foreach [detail in OrderDetails]>>");
            // Each line will display product name and quantity.
            builder.Writeln("- <<[detail.ProductName]>> : <<[detail.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before building the report).
            string templatePath = "OrderReportTemplate.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template tags is "order".
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            string reportPath = "OrderReportResult.docx";
            reportDoc.Save(reportPath);
        }
    }
}
