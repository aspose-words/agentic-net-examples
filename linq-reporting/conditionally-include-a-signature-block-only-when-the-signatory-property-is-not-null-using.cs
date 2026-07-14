using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSignatureExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        // When null the signature block will be omitted.
        public string? Signatory { get; set; }
    }

    public class Program
    {
        // Paths used by the example.
        private static readonly string OutputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        private static readonly string TemplatePath = Path.Combine(OutputDir, "Template.docx");
        private static readonly string ReportWithSignaturePath = Path.Combine(OutputDir, "Report_WithSignature.docx");
        private static readonly string ReportWithoutSignaturePath = Path.Combine(OutputDir, "Report_WithoutSignature.docx");

        public static void Main()
        {
            // Ensure the output folder exists.
            Directory.CreateDirectory(OutputDir);

            // 1. Create the template document with LINQ Reporting tags.
            CreateTemplate();

            // 2. Build a report where the Signatory is provided.
            var modelWithSignature = new ReportModel
            {
                Title = "Monthly Summary",
                Body = "All targets were achieved this month.",
                Signatory = "John Doe"
            };
            BuildReport(modelWithSignature, ReportWithSignaturePath);

            // 3. Build a report where the Signatory is null (signature block omitted).
            var modelWithoutSignature = new ReportModel
            {
                Title = "Monthly Summary",
                Body = "All targets were achieved this month.",
                Signatory = null
            };
            BuildReport(modelWithoutSignature, ReportWithoutSignaturePath);
        }

        private static void CreateTemplate()
        {
            // Create a blank document and a builder to insert content.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Body placeholder.
            builder.Writeln("<<[model.Body]>>");
            builder.Writeln();

            // Conditional block: include the signature line only when Signatory is not null.
            builder.Writeln("<<if [model.Signatory != null]>>");
            builder.Writeln("Signed by: <<[model.Signatory]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            doc.Save(TemplatePath);
        }

        private static void BuildReport(ReportModel model, string outputPath)
        {
            // Load the previously saved template.
            var doc = new Document(TemplatePath);

            // Create the reporting engine.
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
