using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsSummaryExample
{
    // Simple order entity.
    public class Order
    {
        public string Customer { get; set; }
        public decimal Amount { get; set; }

        public Order(string customer, decimal amount)
        {
            Customer = customer;
            Amount = amount;
        }
    }

    // Data source for the reporting engine – contains the aggregated total.
    public class SummaryData
    {
        public decimal Total { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample order data.
            // -----------------------------------------------------------------
            List<Order> orders = new List<Order>
            {
                new Order("Alice", 120.50m),
                new Order("Bob",   75.00m),
                new Order("Carol", 210.30m)
            };

            // -----------------------------------------------------------------
            // 2. Aggregate the order totals using LINQ.
            // -----------------------------------------------------------------
            decimal totalAmount = orders.Sum(o => o.Amount);

            // -----------------------------------------------------------------
            // 3. Load the template document that contains a variable tag
            //    expression like <<[Total]>> where the summary will be placed.
            // -----------------------------------------------------------------
            Document template = new Document("Template.docx"); // assumes the file exists.

            // -----------------------------------------------------------------
            // 4. Build the report by supplying the aggregated total as a data source.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The engine can work with an anonymous object, but using a strongly‑typed
            // class makes the template expression clearer.
            SummaryData data = new SummaryData { Total = totalAmount };

            engine.BuildReport(template, data);

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            template.Save("SummaryReport.docx");
        }
    }
}
