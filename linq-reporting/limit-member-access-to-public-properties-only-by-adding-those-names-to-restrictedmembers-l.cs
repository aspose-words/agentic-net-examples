using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model with public properties.
    public class ReportModel
    {
        // This property is allowed to be accessed from the template.
        public string PublicValue { get; set; } = "Visible";

        // This property should be hidden from the template.
        public string Secret { get; set; } = "Hidden";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any required encodings.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string outputPath = "report.docx";

            // -------------------------------------------------
            // Create a simple Word template with LINQ tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The template references both a public and a restricted member.
            builder.Writeln("Public value: <<[model.PublicValue]>>");
            builder.Writeln("Secret value: <<[model.Secret]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for reporting.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // -------------------------------------------------
            // Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Aspose.Words ReportingEngine does not provide a per‑member
            // restriction list (RestrictedMembers). If you need to block
            // access to an entire type, you can use SetRestrictedTypes,
            // but individual members cannot be hidden this way.
            // For this example we simply omit any restriction.

            // Allow missing members to be treated as empty strings instead of throwing.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = string.Empty;

            // Build the report using the root object name "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
