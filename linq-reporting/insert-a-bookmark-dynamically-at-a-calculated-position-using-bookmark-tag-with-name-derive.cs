using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBookmark
{
    // Simple data model used as the root object for the LINQ Reporting engine.
    public class Order
    {
        // Sample fields.
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = string.Empty;

        // The bookmark name is calculated from other fields.
        // Example: "BM_1", "BM_2", etc.
        public string BookmarkName => $"BM_{Id}";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write some static text and a field that shows the order Id.
            builder.Writeln("Order report");
            builder.Writeln("Order Id: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Insert a bookmark tag whose name is derived from the data model.
            // The content inside the bookmark can be any text; here we use a placeholder.
            builder.Writeln("<<bookmark [order.BookmarkName]>>");
            builder.Writeln("This text is inside the dynamically named bookmark.");
            builder.Writeln("<</bookmark>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order
            {
                Id = 123,
                CustomerName = "John Doe"
                // BookmarkName is calculated automatically.
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport expects the root object name to match the tag prefix ("order").
            engine.BuildReport(report, order, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportWithDynamicBookmark.docx";
            report.Save(outputPath);

            // Inform the user (optional, does not affect the example logic).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
