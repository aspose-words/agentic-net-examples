using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a Stock property.
    public class Product
    {
        public int Stock { get; set; } = 0; // Initialized to avoid nullable warnings.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a conditional block that shows "In stock" when Stock > 0.
            // The root object will be named "product", so we reference it in the expression.
            builder.Writeln("<<if [product.Stock > 0]>>In stock<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare sample data.
            Product sampleProduct = new Product { Stock = 5 }; // Change to 0 to see the block omitted.

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the loaded template, the data source, and the root name "product".
            engine.BuildReport(loadedTemplate, sampleProduct, "product");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
