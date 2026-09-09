using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a Stock property.
    public class Product
    {
        // Initialize to avoid nullable warnings.
        public int Stock { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document and a builder to insert the LINQ Reporting tag.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a conditional block that shows "In stock" when Stock > 0.
            // The root object name will be "product", so we reference product.Stock in the condition.
            builder.Writeln("<<if [product.Stock > 0]>>In stock<</if>>");

            // 2. Prepare the data source.
            var product = new Product { Stock = 5 }; // Change the value to test the condition.

            // 3. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            // The third parameter is the name used in the template to reference the root object.
            engine.BuildReport(doc, product, "product");

            // 4. Save the generated document.
            doc.Save("Report_Output.docx");
        }
    }
}
