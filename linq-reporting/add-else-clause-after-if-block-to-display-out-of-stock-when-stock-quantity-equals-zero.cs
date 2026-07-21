using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing a product.
    public class Product
    {
        // Name of the product.
        public string Name { get; set; } = string.Empty;

        // Quantity in stock.
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert product name placeholder.
            builder.Writeln("Product: <<[product.Name]>>");

            // Insert quantity placeholder with an if‑else logic.
            // If Quantity > 0, display the numeric value; otherwise display "Out of stock".
            builder.Writeln(
                "Quantity: " +
                "<<if [product.Quantity > 0]>>" +
                "<<[product.Quantity]>>" +
                "<</if>>" +
                "<<if [product.Quantity == 0]>>Out of stock<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // Step 3: Prepare sample data.
            Product sampleProduct = new Product
            {
                Name = "Widget",
                Quantity = 0 // Zero quantity triggers the else clause.
            };

            // Step 4: Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "product".
            engine.BuildReport(reportDoc, sampleProduct, "product");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
