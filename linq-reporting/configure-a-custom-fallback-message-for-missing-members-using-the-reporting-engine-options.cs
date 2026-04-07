using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public Customer Customer { get; set; } = new Customer();
    }

    public class Customer
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders for the template and the generated report.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "Template.docx");
            string outputPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document that contains a reference to a missing member.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Existing member – will be replaced with the actual value.
            builder.Writeln("Customer name: <<[model.Customer.Name]>>");

            // Missing member – the engine will use the custom fallback message.
            builder.Writeln("Missing field: <<[model.MissingObject.SomeProperty]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and configure the ReportingEngine.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Create the data source object.
            ReportModel model = new ReportModel();

            // Configure the engine to allow missing members and provide a custom message.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "[Member not found]";

            // Build the report. The root object is referenced in the template as "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);

            // Inform the user where the files are located.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}
