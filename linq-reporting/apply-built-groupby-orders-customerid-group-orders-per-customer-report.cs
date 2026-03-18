using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple POCO representing an order.
    public class Order
    {
        public int CustomerId { get; set; }
        public string Product   { get; set; }
        public int    Quantity  { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare the data source – a list of orders that will be grouped
            //    by CustomerId in the report.
            // -----------------------------------------------------------------
            var orders = new List<Order>
            {
                new Order { CustomerId = 1, Product = "Apple",  Quantity = 5 },
                new Order { CustomerId = 2, Product = "Banana", Quantity = 3 },
                new Order { CustomerId = 1, Product = "Orange", Quantity = 2 },
                new Order { CustomerId = 3, Product = "Grape",  Quantity = 7 },
                new Order { CustomerId = 2, Product = "Pear",   Quantity = 4 }
            };

            // -----------------------------------------------------------------
            // 2. Load the template document that contains the ReportingEngine
            //    tags for grouping. The template must include something like:
            //    
            //    <<groupby [orders] CustomerId>>
            //        Customer ID: <<[orders.CustomerId]>>
            //        <<foreach [orders]>>
            //            - <<[Product]>> (Qty: <<[Quantity]>>)
            //        <<endforeach>>
            //    <<endgroupby>>
            //    
            //    The actual template file (Template.docx) is assumed to exist.
            // -----------------------------------------------------------------
            Document doc = new Document("Template.docx");

            // -----------------------------------------------------------------
            // 3. Build the report. The anonymous object supplies a property named
            //    "orders" that the template can reference. The third argument
            //    ("data") is the name used inside the template to refer to the
            //    whole data source object (optional – can be null or empty).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new { orders }, "data");

            // -----------------------------------------------------------------
            // 4. Save the populated document.
            // -----------------------------------------------------------------
            doc.Save("Report.docx");
        }
    }
}
