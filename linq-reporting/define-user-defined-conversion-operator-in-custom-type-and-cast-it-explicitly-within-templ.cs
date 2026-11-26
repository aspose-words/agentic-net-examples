using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConversionExample
{
    // Custom type with an explicit conversion operator to decimal.
    public class Price
    {
        public decimal Value { get; set; }

        public Price(decimal value) => Value = value;

        // Explicit conversion from Price to decimal.
        public static explicit operator decimal(Price p) => p.Value;

        // Optional: provide a string representation.
        public override string ToString() => Value.ToString("C");
    }

    // Model class used as the root data source for the report.
    public class Order
    {
        public int Id { get; set; }

        // Initialize to avoid nullable warnings.
        public Price Price { get; set; } = null!;
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert simple text and LINQ Reporting tags.
            builder.Writeln("Order ID: <<[order.Id]>>");
            // Explicitly cast the custom type to decimal within the expression.
            builder.Writeln("Order Price: <<[(decimal)order.Price]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and prepare the data source.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            Order order = new Order
            {
                Id = 1001,
                Price = new Price(199.99m)
            };

            // -------------------------------------------------
            // Build the report using Aspose.Words LINQ Reporting.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
