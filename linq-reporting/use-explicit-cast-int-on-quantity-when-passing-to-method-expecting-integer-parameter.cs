using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used as the root object for the report.
    public class Order
    {
        // Product name.
        public string Product { get; set; } = string.Empty;

        // Quantity stored as double to demonstrate explicit casting.
        public double Quantity { get; set; }

        // Method that expects an integer parameter.
        public string GetQuantityString(int qty)
        {
            return $"Qty: {qty}";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            // The expression casts Quantity (double) to int before passing it to GetQuantityString.
            builder.Writeln("Product: <<[order.Product]>>");
            builder.Writeln("Quantity: <<[order.GetQuantityString((int)order.Quantity)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // Sample data source.
            var order = new Order
            {
                Product = "Apple",
                Quantity = 7.9 // Will be cast to 7 in the template.
            };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
