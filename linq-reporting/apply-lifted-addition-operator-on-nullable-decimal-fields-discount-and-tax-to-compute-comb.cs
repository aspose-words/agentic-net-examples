using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class Order
    {
        // Nullable decimal fields.
        public decimal? Discount { get; set; } = 5.5m;   // Example value; can be null.
        public decimal? Tax { get; set; } = null;      // Example value; can be null.

        // Combined value using the lifted addition operator.
        public decimal? Combined => Discount + Tax;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple paragraph with a LINQ Reporting tag that references the Combined property.
            builder.Writeln("Combined amount: <<[order.Combined]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template back (required by the lifecycle rule).
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            Order order = new Order(); // Discount = 5.5, Tax = null → Combined = 5.5

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "order", so we pass it explicitly.
            engine.BuildReport(doc, order, "order");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
