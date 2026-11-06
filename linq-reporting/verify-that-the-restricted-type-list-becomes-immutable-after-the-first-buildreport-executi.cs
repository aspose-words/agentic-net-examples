using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingImmutableRestrictedTypes
{
    // Simple data model used in the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Create a template document that contains a LINQ Reporting tag.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<[model.Name]>>");
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document doc = new Document(templatePath);

            // Prepare the root data object.
            ReportModel model = new ReportModel { Name = "Aspose Test" };

            // -----------------------------------------------------------------
            // Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The ReportingEngine class provides a RestrictedTypes collection that
            // becomes read‑only after the first BuildReport call.  The property may be
            // non‑public in some library versions, so we obtain it via reflection.
            PropertyInfo? restrictedProp = typeof(ReportingEngine).GetProperty(
                "RestrictedTypes",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (restrictedProp == null)
            {
                Console.WriteLine("RestrictedTypes property not found on ReportingEngine.");
                return;
            }

            // The property is expected to be an IList<Type>.
            if (restrictedProp.GetValue(engine) is not IList<Type> restrictedTypes)
            {
                Console.WriteLine("Unable to retrieve RestrictedTypes collection.");
                return;
            }

            // Add a type before building the report.
            restrictedTypes.Add(typeof(ReportModel));

            // Build the report.
            bool success = engine.BuildReport(doc, model, "model");
            Console.WriteLine($"BuildReport success: {success}");

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "result.docx");
            doc.Save(resultPath);
            Console.WriteLine($"Report saved to: {resultPath}");

            // -----------------------------------------------------------------
            // Verify that the RestrictedTypes collection becomes immutable after
            // the first BuildReport execution.
            // -----------------------------------------------------------------
            try
            {
                restrictedTypes.Add(typeof(string));
                Console.WriteLine("RestrictedTypes is still mutable (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"RestrictedTypes immutable as expected: {ex.GetType().Name}");
            }
        }
    }
}
