using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model with a property named "new". The property name is escaped with @ in C# code.
    public class SampleModel
    {
        // The property name is "new", which is a C# reserved keyword.
        // It must be prefixed with @ when referenced in C# code.
        public string @new { get; set; } = "Escaped keyword value";

        public string Name { get; set; } = "Sample Model";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // ---------- Create the template document ----------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Use LINQ Reporting tags. The property named "new" is accessed as model.new (no @ in the tag).
            builder.Writeln("Escaped property value: <<[model.new]>>");
            builder.Writeln("Model name: <<[model.Name]>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // ---------- Load the template and build the report ----------
            Document loadedTemplate = new Document(templatePath);

            // Create the data model instance.
            SampleModel model = new SampleModel();

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
