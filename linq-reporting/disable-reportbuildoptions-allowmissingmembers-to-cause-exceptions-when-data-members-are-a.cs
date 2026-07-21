using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model without the property referenced in the template.
    public class ReportModel
    {
        // Existing property to avoid nullable warnings.
        public string Existing { get; set; } = "Existing value";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the template and the output document.
            const string templatePath = "Template.docx";
            const string outputPath = "Result.docx";

            // -------------------------------------------------
            // Create a template document with a missing member tag.
            // -------------------------------------------------
            var builder = new DocumentBuilder();
            // The tag references a property that does NOT exist in ReportModel.
            builder.Writeln("<<[MissingObject.Name]>>");
            // Save the template to disk.
            builder.Document.Save(templatePath);

            // Load the template back for reporting.
            var doc = new Document(templatePath);

            // Prepare a data source that lacks the referenced member.
            var model = new ReportModel();

            // Configure the reporting engine without AllowMissingMembers.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // Explicitly disable all special options.

            try
            {
                // Attempt to build the report. This should throw because the member is missing.
                engine.BuildReport(doc, model, "model");
                // If no exception occurs, save the (unexpected) result.
                doc.Save(outputPath);
                Console.WriteLine("Report built successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown due to the missing member.
                Console.WriteLine("Exception caught as expected:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
