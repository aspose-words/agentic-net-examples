using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        // Non‑nullable reference type initialized to avoid warnings.
        public string CustomerName { get; set; } = string.Empty;

        // Empty string will cause the corresponding paragraph to become empty.
        public string Note { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that will always contain a value.
            builder.Writeln("<<[order.CustomerName]>>");

            // Paragraph that may become empty after the tag is replaced.
            builder.Writeln("<<[order.Note]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data source.
            // -------------------------------------------------
            Order order = new Order
            {
                CustomerName = "John Doe",
                Note = string.Empty // Intentionally empty to trigger paragraph removal.
            };

            // -------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // -------------------------------------------------
            // 5. Measure memory before and after BuildReport.
            // -------------------------------------------------
            // Force a full garbage collection to get a clean baseline.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);

            // Build the report. The root object name must match the tag prefix ("order").
            engine.BuildReport(reportDoc, order, "order");

            // Force another collection before measuring the after state.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
            long memoryConsumed = memoryAfter - memoryBefore;

            // -------------------------------------------------
            // 6. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);

            // -------------------------------------------------
            // 7. Output memory consumption.
            // -------------------------------------------------
            Console.WriteLine($"Memory before BuildReport: {memoryBefore:N0} bytes");
            Console.WriteLine($"Memory after  BuildReport: {memoryAfter:N0} bytes");
            Console.WriteLine($"Memory consumed by BuildReport: {memoryConsumed:N0} bytes");
        }
    }
}
