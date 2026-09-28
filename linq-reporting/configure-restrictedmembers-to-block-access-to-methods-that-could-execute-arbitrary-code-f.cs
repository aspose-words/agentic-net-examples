using System;
using System.Collections;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedMembersExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Safe property that will be displayed.
        public string Name { get; set; } = string.Empty;

        // Method that could execute arbitrary code – we will block it.
        public string DangerousMethod()
        {
            // In a real scenario this could perform unsafe actions.
            return "Executed dangerous code!";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document with LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln("Attempted Dangerous Call: <<[model.DangerousMethod]>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            templateDoc.Save(templatePath);

            // Load the template back for reporting.
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new()
            {
                Name = "John Doe"
            };

            // Configure the ReportingEngine.
            ReportingEngine engine = new();

            // Enable inline error messages so blocked members are shown in the output.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Block the DangerousMethod to prevent its execution.
            // Use reflection to access the RestrictedMembers collection if it exists.
            var restrictedProp = typeof(ReportingEngine).GetProperty("RestrictedMembers");
            if (restrictedProp != null)
            {
                if (restrictedProp.GetValue(engine) is IList list)
                {
                    list.Add(nameof(ReportModel.DangerousMethod));
                }
            }

            // Build the report.
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string outputPath = "output.docx";
            loadedTemplate.Save(outputPath);

            // Write simple status to the console.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
