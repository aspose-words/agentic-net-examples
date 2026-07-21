using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLiftedAndExample
{
    // Data model used by the LINQ Reporting template.
    public class FeatureModel
    {
        // Nullable booleans to demonstrate lifted logical operators.
        public bool? IsActive { get; set; } = false;
        public bool? HasLicense { get; set; } = false;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Use explicit comparisons to avoid applying && directly to nullable booleans.
            // The condition evaluates to true only when both values are true.
            builder.Writeln("<<if [model.IsActive == true && model.HasLicense == true]>>Feature is AVAILABLE<</if>>");
            // Show the alternative message when the condition is not met.
            builder.Writeln("<<if [!(model.IsActive == true && model.HasLicense == true)]>>Feature is NOT AVAILABLE<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var data = new FeatureModel
            {
                IsActive = true,   // Both conditions are true → feature should be AVAILABLE.
                HasLicense = true
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, data, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
