using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple order class with an amount.
    public class Order
    {
        public decimal Amount { get; set; } = 0;
    }

    // Wrapper model that holds a collection of orders and provides calculated fields.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();

        // Total sales calculated from the order amounts.
        public decimal TotalSales => Orders.Sum(o => o.Amount);

        // Average order value; returns 0 when there are no orders to avoid division by zero.
        public decimal AverageOrderValue => Orders.Any() ? Orders.Average(o => o.Amount) : 0;
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------- Data bootstrap --------------------
            var model = new ReportModel();
            model.Orders.AddRange(new[]
            {
                new Order { Amount = 120.50m },
                new Order { Amount = 75.00m },
                new Order { Amount = 200.00m }
            });

            // -------------------- Template creation --------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Sales Summary");
            // Expression tags that will be replaced by the reporting engine.
            builder.Writeln("Total Sales: <<[model.TotalSales]>>");
            builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

            const string templatePath = "SalesSummaryTemplate.docx";
            doc.Save(templatePath); // Save the template before building the report.

            // -------------------- Report generation --------------------
            var template = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            const string outputPath = "SalesSummaryReport.docx";
            template.Save(outputPath); // Final report with calculated values.
        }
    }
}
