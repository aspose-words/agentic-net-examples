using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Product
    {
        // Name of the product – initialized to avoid nullable warnings.
        public string Name { get; set; } = "Sample Product";

        // Stock quantity.
        public int Stock { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Insert LINQ Reporting tags.
            //    Show the product name.
            builder.Writeln("Product: <<[product.Name]>>");

            //    Conditional block – displays "In stock" only when Stock > 0.
            builder.Writeln("<<if [product.Stock > 0]>>In stock<</if>>");

            // 3. Prepare the data source.
            Product product = new Product
            {
                Name = "Aspose.Words Book",
                Stock = 5 // Change to 0 to see the condition evaluate to false.
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name ("product") must match the tag prefix used in the template.
            engine.BuildReport(template, product, "product");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
