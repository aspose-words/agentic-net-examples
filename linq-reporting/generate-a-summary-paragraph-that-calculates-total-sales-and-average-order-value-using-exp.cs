using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingSummary
{
    // Data model representing a single order.
    public class Order
    {
        public decimal Amount { get; set; } = 0m;
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();

        // Calculated total sales.
        public decimal TotalSales => Orders.Sum(o => o.Amount);

        // Calculated average order value.
        public decimal AverageOrderValue => Orders.Any() ? Orders.Average(o => o.Amount) : 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("Sales Summary Report");
            builder.Writeln();

            // Insert expression tags for total sales and average order value.
            builder.Writeln("Total Sales: <<[model.TotalSales]>>");
            builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template and prepare sample data.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data: a few orders with different amounts.
            ReportModel model = new()
            {
                Orders = new List<Order>
                {
                    new Order { Amount = 120.50m },
                    new Order { Amount = 75.00m },
                    new Order { Amount = 200.00m },
                    new Order { Amount = 50.25m }
                }
            };

            // -------------------------------------------------
            // Step 3: Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // Step 4: Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
