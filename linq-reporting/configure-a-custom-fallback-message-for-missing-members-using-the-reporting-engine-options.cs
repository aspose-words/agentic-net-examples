using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // File paths for the template and the generated report
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // 1. Create a template document that contains a tag referencing a missing member
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // The expression tries to access Customer.NonExistent, which does not exist
            builder.Writeln("Customer name: <<[customer.NonExistent]>>");
            templateDoc.Save(templatePath);

            // 2. Load the template back from disk (simulating a separate load step)
            Document doc = new Document(templatePath);

            // 3. Prepare a data model that does NOT contain the missing member
            var model = new ReportModel
            {
                customer = new Customer { Name = "John Doe" }
            };

            // 4. Configure the ReportingEngine
            ReportingEngine engine = new ReportingEngine();
            // Allow missing members so the engine substitutes them instead of throwing
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Custom fallback message that will appear in place of the missing member
            engine.MissingMemberMessage = "[missing]";

            // 5. Build the report using the model; the root name in the template is "model"
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report
            doc.Save(outputPath);
        }
    }

    // Public wrapper class used as the root data source for the report
    public class ReportModel
    {
        public Customer customer { get; set; } = new Customer();
    }

    // Simple data model class; it intentionally lacks the "NonExistent" property
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
    }
}
