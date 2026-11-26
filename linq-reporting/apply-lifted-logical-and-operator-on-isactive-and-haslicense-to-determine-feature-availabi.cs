using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLiftedAndExample
{
    // Data model with nullable booleans.
    public class FeatureModel
    {
        // Nullable booleans to demonstrate lifted logical AND.
        public bool? IsActive { get; set; } = false;
        public bool? HasLicense { get; set; } = false;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new FeatureModel
            {
                IsActive = true,
                HasLicense = true
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Feature Availability Report");
            builder.Writeln("<<if [model.IsActive & model.HasLicense]>>Feature is AVAILABLE<</if>>");
            builder.Writeln("<<if [!(model.IsActive & model.HasLicense)]>>Feature is UNAVAILABLE<</if>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
