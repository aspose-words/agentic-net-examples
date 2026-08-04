using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Product
    {
        // Stock quantity – non‑nullable to avoid warnings.
        public int Stock { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Product stock status:");
            // Conditional block: show "In stock" only when Stock > 0.
            builder.Writeln("<<if [model.Stock > 0]>>In stock<</if>>");
            // Optional: show "Out of stock" when the condition is false.
            builder.Writeln("<<if [model.Stock <= 0]>>Out of stock<</if>>");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Product product = new Product { Stock = 5 }; // Change the value to test different outcomes.

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, product, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
