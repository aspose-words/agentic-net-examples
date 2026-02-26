using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model that will be used as the data source for the report.
    public class Order
    {
        public int Id { get; set; }
        public string Customer { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    // Wrapper class that exposes a LINQ query as a property.
    // The template can reference this property directly.
    public class ReportData
    {
        // Original collection of orders.
        public List<Order> Orders { get; set; }

        // Example of a LINQ expression using a lambda function.
        // Returns only orders with an amount greater than the specified threshold.
        public IEnumerable<Order> HighValueOrders => Orders
            .Where(o => o.Amount > 1000m)          // Lambda filter
            .OrderByDescending(o => o.Amount);    // Lambda ordering
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOCM template that contains Reporting Engine tags.
            //    The template should have tags like <<foreach [ds.HighValueOrders]>>
            //    <<[Id]>> <<[Customer]>> <<[Amount]>> <</foreach>>
            Document doc = new Document("Template.docm");

            // 2. Prepare the data source.
            var data = new ReportData
            {
                Orders = new List<Order>
                {
                    new Order { Id = 1, Customer = "Alpha Co.", Amount = 750m,  Date = new DateTime(2023, 1, 15) },
                    new Order { Id = 2, Customer = "Beta Ltd.", Amount = 1250m, Date = new DateTime(2023, 2, 3) },
                    new Order { Id = 3, Customer = "Gamma Inc.", Amount = 2100m, Date = new DateTime(2023, 3, 22) },
                    new Order { Id = 4, Customer = "Delta LLC", Amount = 500m,  Date = new DateTime(2023, 4, 10) }
                }
            };

            // 3. Build the report using the ReportingEngine.
            //    The data source name "ds" matches the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // 4. Save the populated document.
            doc.Save("Report.docx");
        }
    }
}
