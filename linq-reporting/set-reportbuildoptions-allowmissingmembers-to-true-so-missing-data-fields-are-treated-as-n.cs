using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingAllowMissingMembers
{
    // Simple data model with only the Name property.
    public class Model
    {
        public string Name { get; set; } = "John Doe";
        // Note: No MissingField property – it will be missing in the template.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a temporary folder for the example files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Customer: <<[model.Name]>>");
            builder.Writeln("Missing field (should be empty): <<[model.MissingField]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real-world scenario).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            DocumentBuilder reportBuilder = new DocumentBuilder(doc); // Not strictly needed but kept for consistency.

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            Model model = new Model();

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to allow missing members.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Build the report. Missing fields will be treated as null (empty output).
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);

            // The example finishes here. No interactive prompts are used.
        }
    }
}
