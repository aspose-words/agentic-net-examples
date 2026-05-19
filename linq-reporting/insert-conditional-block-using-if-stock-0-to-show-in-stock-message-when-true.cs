using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

// Ensure code page support (required by Aspose.Words for some encodings)
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a non‑nullable Stock property
    public class Product
    {
        public int Stock { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document that will serve as the template
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Insert a conditional LINQ Reporting tag:
            //    Show "In stock" only when the Stock value is greater than 0
            builder.Writeln("<<if [product.Stock > 0]>>In stock<</if>>");

            // 3. Prepare sample data
            Product product = new Product { Stock = 5 }; // Change to 0 to see no output

            // 4. Build the report using the template and the data source
            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template is "product"
            engine.BuildReport(template, product, "product");

            // 5. Save the generated document
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            template.Save(outputPath);
        }
    }
}
