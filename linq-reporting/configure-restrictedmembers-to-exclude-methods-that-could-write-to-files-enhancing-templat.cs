using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public string CustomerName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the data model.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Configure restricted members to block file‑writing APIs.
            // -----------------------------------------------------------------
            // Restrict types that expose methods capable of writing to the file system.
            ReportingEngine.SetRestrictedTypes(
                typeof(System.IO.File),
                typeof(System.IO.FileInfo),
                typeof(System.IO.Directory),
                typeof(System.IO.StreamWriter),
                typeof(System.IO.BinaryWriter));

            // -----------------------------------------------------------------
            // 3. Build the report using the template and a sample data object.
            // -----------------------------------------------------------------
            // Load the template back (simulating a real‑world scenario where the
            // template might be stored separately).
            Document loadedTemplate = new Document(templatePath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Prepare sample data.
            Order sampleOrder = new Order { CustomerName = "John Doe" };

            // Build the report. The root object name ("order") must match the tag.
            engine.BuildReport(loadedTemplate, sampleOrder, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
