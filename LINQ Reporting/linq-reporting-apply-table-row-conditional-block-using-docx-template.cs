using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data model for the table rows.
    public class Order
    {
        public string Item { get; set; }
        public int Quantity { get; set; }

        // This property will be used in the template to conditionally display the row.
        public bool ShowRow => Quantity > 0;
    }

    // Wrapper class that contains the collection used by the template.
    public class ReportData
    {
        public List<Order> Orders { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains a table with conditional tags.
            // The template should have a row like:
            // <<if [Orders[i].ShowRow]>>
            //   <<[Orders[i].Item]>>
            //   <<[Orders[i].Quantity]>>
            // <<endif>>
            Document doc = new Document("TemplateWithConditionalRow.docx");

            // Prepare the data source.
            var data = new ReportData
            {
                Orders = new List<Order>
                {
                    new Order { Item = "Apples",  Quantity = 5 },
                    new Order { Item = "Bananas", Quantity = 0 }, // This row will be omitted.
                    new Order { Item = "Oranges", Quantity = 12 }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object.
            // The third parameter (optional) gives a name to reference the source inside the template.
            engine.BuildReport(doc, data, "data");

            // Save the populated document.
            doc.Save("ReportWithConditionalRows.docx");
        }
    }
}
