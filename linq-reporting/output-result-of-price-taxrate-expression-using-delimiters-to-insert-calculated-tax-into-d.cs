using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model that will be used as the root object for the report.
    public class Order
    {
        // Price of the item.
        public double Price { get; set; } = 0;

        // Tax rate expressed as a decimal (e.g., 0.20 for 20%).
        public double TaxRate { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and a builder to compose the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Write static text and a LINQ Reporting tag that calculates price * taxRate.
            // The tag uses the root name "order" which will be supplied when building the report.
            builder.Writeln("Price: $<<[order.Price]>>");
            builder.Writeln("Tax Rate: <<[order.TaxRate]>> (as decimal)");
            builder.Writeln("Calculated Tax: $<<[order.Price * order.TaxRate]>>");

            // 3. Prepare sample data.
            Order sampleOrder = new Order
            {
                Price = 199.99,
                TaxRate = 0.07 // 7% tax
            };

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The template references the root object as "order", so we pass the name accordingly.
            engine.BuildReport(doc, sampleOrder, "order");

            // 5. Save the generated document.
            doc.Save("CalculatedTaxReport.docx");
        }
    }
}
