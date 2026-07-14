using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLiftedAndExample
{
    // Model class used as the data source for the report.
    public class FeatureModel
    {
        // Name of the feature.
        public string Feature { get; set; } = "Premium Feature";

        // Nullable boolean indicating whether the feature is active.
        public bool? IsActive { get; set; }

        // Nullable boolean indicating whether the user has a license for the feature.
        public bool? HasLicense { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a title.
            builder.Writeln("Feature Availability Report");
            builder.Writeln();

            // Insert the feature name.
            builder.Writeln("Feature: <<[model.Feature]>>");
            builder.Writeln();

            // Use a lifted logical AND (&&) on the nullable booleans.
            // The expression evaluates to true only when both operands are true.
            // If either operand is null, it is treated as false via the null‑coalescing operator.
            builder.Writeln("<<if [(model.IsActive ?? false) && (model.HasLicense ?? false)]>>Status: Available<</if>>");
            builder.Writeln("<<if [!((model.IsActive ?? false) && (model.HasLicense ?? false))]>>Status: Not Available<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            FeatureModel model = new FeatureModel
            {
                // Both conditions are true → feature is available.
                IsActive = true,
                HasLicense = true
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
