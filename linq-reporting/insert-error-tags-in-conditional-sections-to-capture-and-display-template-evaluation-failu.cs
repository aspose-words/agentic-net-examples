using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingErrorDemo
{
    // Sample data model
    public class ReportModel
    {
        public string Name { get; set; } = "Sample Report";
        // Intentionally no property named MissingProp to trigger an evaluation error
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Simple text with a property reference
            builder.Writeln("Report Title: <<[model.Name]>>");
            builder.Writeln();

            // Conditional section that references a non‑existent property.
            // The <<error>> tag will display the evaluation failure.
            builder.Writeln("<<if [model.MissingProp]>>");
            builder.Writeln("This text is inside the true branch.");
            builder.Writeln("<<error>>"); // Capture and display the error
            builder.Writeln("<</if>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // Prepare the root data object
            var model = new ReportModel();

            // Configure the reporting engine to emit inline error messages
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report
            doc.Save(outputPath);

            // Output the result status
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
