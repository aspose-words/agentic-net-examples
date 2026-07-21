using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with only one property.
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        // Note: Age property is intentionally omitted to demonstrate missing member handling.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create a template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags. The second tag references a missing member (Age).
            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln("Customer Age: <<[model.Age]>>"); // Age does not exist in the model.

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Treat missing members as null (or empty) values.
                Options = ReportBuildOptions.AllowMissingMembers,
                // Optional: custom text to display for missing members.
                MissingMemberMessage = "N/A"
            };

            // -------------------------------------------------
            // 4. Prepare the data source.
            // -------------------------------------------------
            Customer model = new Customer { Name = "John Doe" };

            // -------------------------------------------------
            // 5. Build the report.
            // -------------------------------------------------
            // The root object name used in the template tags is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 6. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
