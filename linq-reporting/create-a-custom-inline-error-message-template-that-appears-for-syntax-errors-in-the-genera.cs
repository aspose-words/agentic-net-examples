using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessageExample
{
    // Simple data model for the report.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template with a deliberate syntax error.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the model's Name.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Deliberate syntax error – missing closing bracket.
            builder.Writeln("This line contains an error: <<[model.Missing>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            Model model = new Model();

            // -----------------------------------------------------------------
            // 4. Build the report with inline error messages enabled.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            bool success = engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Output the result status (optional, not interactive).
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
